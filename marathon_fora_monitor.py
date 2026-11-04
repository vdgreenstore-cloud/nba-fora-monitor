#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Мониторинг фор (гандикапов) на матчи NBA у букмекера Marathonbet.

Что делает:
  1. Скачивает страницу https://www.marathonbet.ru/su/betting/Basketball/NBA
  2. По каждой игре извлекает главную фору и коэффициенты на неё
  3. Сравнивает с предыдущим запуском и построчно записывает по каждой игре,
     изменилась линия или нет:
        - в читаемый лог   log.txt
        - в историю        history.csv (открывается в Excel)
  4. Текущий снимок хранит в state.json

Режимы:
  python3 marathon_fora_monitor.py --once                  одна проверка
  python3 marathon_fora_monitor.py --loop                  вечный цикл (каждые 30 мин)
  python3 marathon_fora_monitor.py --loop --interval 600   свой интервал в секундах

Зависимости: только стандартная библиотека Python 3.8+.
"""

import argparse
import csv
import gzip
import html
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

DEFAULT_URL = "https://www.marathonbet.ru/su/betting/Basketball/NBA"
DEFAULT_INTERVAL = 1800  # 30 минут
DEFAULT_TIMEOUT = 40
DEFAULT_RETRIES = 3

MSK = timezone(timedelta(hours=3))  # московское время, без перехода на летнее

CSV_COLUMNS = [
    "время проверки", "id матча", "матч", "начало (МСК)",
    "фора 1", "кэф 1", "фора 2", "кэф 2", "статус",
]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


# ----------------------------------------------------------------------------
# Скачивание страницы
# ----------------------------------------------------------------------------

def fetch_page(url, timeout=DEFAULT_TIMEOUT, retries=DEFAULT_RETRIES):
    """Скачивает HTML страницы. Возвращает текст или бросает RuntimeError."""
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "ru,en;q=0.8",
            })
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                if resp.status != 200:
                    raise RuntimeError("HTTP {}".format(resp.status))
                raw = resp.read()
                if resp.headers.get("Content-Encoding") == "gzip":
                    raw = gzip.decompress(raw)
                return raw.decode("utf-8", errors="replace")
        except Exception as err:  # noqa: BLE001 - логируем любую ошибку сети
            last_err = err
            if attempt < retries:
                time.sleep(3 * attempt)
    raise RuntimeError("не удалось скачать страницу после {} попыток: {}".format(retries, last_err))


# ----------------------------------------------------------------------------
# Разбор страницы
# ----------------------------------------------------------------------------

def _attr(tag, name):
    m = re.search(name + r'="([^"]*)"', tag)
    return html.unescape(m.group(1)) if m else None


def parse_events(page_html):
    """Извлекает список матчей с форой из HTML страницы Marathonbet."""
    events = []
    # Каждый матч — блок <div class="bg coupon-row" data-event-...> ... </div>
    starts = [m.start() for m in re.finditer(r'<div class="bg coupon-row"', page_html)]
    for i, start in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(page_html)
        chunk = page_html[start:end]

        head = chunk[:chunk.find(">") + 1]  # только открывающий тег блока
        event_id = _attr(head, r"data-event-eventId")
        match_name = _attr(head, r"data-event-name")
        live = 'data-live="true"' in head

        if not event_id or not match_name:
            continue

        # Названия команд
        teams = re.findall(r'<span data-member-link="true">([^<]+)</span>', chunk)
        teams = [t.strip() for t in teams]
        team1 = teams[0] if len(teams) > 0 else ""
        team2 = teams[1] if len(teams) > 1 else ""

        # Время начала (первое вхождение в блоке)
        m_date = re.search(r'<div class="date-wrapper">\s*([^<]+?)\s*</div>', chunk)
        start_time = html.unescape(m_date.group(1)).strip() if m_date else ""

        # Ячейки форы: первая — Фора1, вторая — Фора2
        hcap_cells = re.findall(
            r'<td[^>]*data-market-type="HANDICAP"[^>]*>(.*?)</td>', chunk, re.S)
        f1 = k1 = f2 = k2 = None
        if len(hcap_cells) >= 1:
            f1, k1 = _line_price(hcap_cells[0])
        if len(hcap_cells) >= 2:
            f2, k2 = _line_price(hcap_cells[1])

        events.append({
            "id": event_id,
            "match": match_name,
            "team1": team1,
            "team2": team2,
            "start": start_time,
            "live": live,
            "f1": norm_line(f1),
            "k1": norm_price(k1),
            "f2": norm_line(f2),
            "k2": norm_price(k2),
        })
    return events


def _line_price(cell_html):
    """Из ячейки форы достаёт значение линии и коэффициент."""
    m_line = re.search(r'\(\s*([-+\d.]+)\s*\)', cell_html)
    m_price = re.search(r'data-selection-price="([\d.]+)"', cell_html)
    line = m_line.group(1) if m_line else None
    price = m_price.group(1) if m_price else None
    return line, price


def norm_line(value):
    """Нормализует линию: '-3.50' -> '-3.5', '3' -> '+3.0'; None -> None."""
    if value is None:
        return None
    try:
        v = float(value)
    except ValueError:
        return None
    return "{:+.1f}".format(v)


def norm_price(value):
    if value is None:
        return None
    try:
        return "{:.2f}".format(float(value))
    except ValueError:
        return None


# ----------------------------------------------------------------------------
# Хранение: state.json, history.csv, log.txt
# ----------------------------------------------------------------------------

def load_state(path):
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_state(path, state):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def append_csv(path, row):
    is_new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        if is_new:
            writer.writerow(CSV_COLUMNS)
        writer.writerow(row)


def append_log(path, text):
    with open(path, "a", encoding="utf-8") as f:
        f.write(text + "\n")


# ----------------------------------------------------------------------------
# Один цикл проверки
# ----------------------------------------------------------------------------

def snapshot_tuple(ev):
    return (ev["f1"], ev["k1"], ev["f2"], ev["k2"])


def fmt_fora(ev):
    """Красивое описание текущей линии для лога."""
    if ev["f1"] is None and ev["f2"] is None:
        return "фора не выставлена"
    parts = []
    name1 = ev["team1"] or "команда 1"
    name2 = ev["team2"] or "команда 2"
    parts.append("{}: {} ({})".format(
        name1, ev["f1"] if ev["f1"] is not None else "—",
        ev["k1"] if ev["k1"] is not None else "—"))
    parts.append("{}: {} ({})".format(
        name2, ev["f2"] if ev["f2"] is not None else "—",
        ev["k2"] if ev["k2"] is not None else "—"))
    return "фора " + ", ".join(parts)


def fmt_changes(prev, ev):
    """Перечисляет, что именно изменилось по сравнению с прошлым снимком."""
    labels = [
        ("фора 1", 0, 1, lambda v: v if v is not None else "—"),
        ("кэф 1", 1, 0, lambda v: v if v is not None else "—"),
        ("фора 2", 2, 3, lambda v: v if v is not None else "—"),
        ("кэф 2", 3, 0, lambda v: v if v is not None else "—"),
    ]
    diffs = []
    for label, cur_idx, _, fmt in labels:
        cur_val = snapshot_tuple(ev)[cur_idx]
        old_val = prev[cur_idx]
        if cur_val != old_val:
            diffs.append("{}: {} → {}".format(label, fmt(old_val) if old_val else "—",
                                              fmt(cur_val)))
    return "; ".join(diffs)


def run_once(data_dir, url, timeout=DEFAULT_TIMEOUT, retries=DEFAULT_RETRIES, quiet=False):
    """Одна проверка. Возвращает словарь со сводкой запуска."""
    os.makedirs(data_dir, exist_ok=True)
    state_path = os.path.join(data_dir, "state.json")
    csv_path = os.path.join(data_dir, "history.csv")
    log_path = os.path.join(data_dir, "log.txt")

    now = datetime.now(MSK)
    ts = now.strftime("%d.%m.%Y %H:%M:%S")

    try:
        page_html = fetch_page(url, timeout=timeout, retries=retries)
        events = parse_events(page_html)
        if not events:
            raise RuntimeError("на странице не найдено ни одного матча — "
                               "возможно, сайт изменил разметку")
    except Exception as err:  # noqa: BLE001
        msg = "[!] {} — ОШИБКА: {}".format(ts, err)
        append_log(log_path, msg)
        print(msg, file=sys.stderr)
        return {"ok": False, "error": str(err)}

    prev_state = load_state(state_path)  # id -> [f1, k1, f2, k2]

    stats = {"total": len(events), "new": 0, "changed": 0, "same": 0, "noline": 0}

    log_lines = ["=== {} МСК | матчей на странице: {} ===".format(ts, len(events))]
    for ev in events:
        cur = snapshot_tuple(ev)
        prev = prev_state.get(ev["id"])

        if prev is None:
            status = "НОВЫЙ МАТЧ"
            stats["new"] += 1
            detail = fmt_fora(ev)
        elif cur == tuple(prev):
            status = "БЕЗ ИЗМЕНЕНИЙ"
            stats["same"] += 1
            detail = fmt_fora(ev)
        else:
            status = "ИЗМЕНЕНИЕ"
            stats["changed"] += 1
            detail = "было [{} ({}) / {} ({})] → стало [{} ({}) / {} ({})]; ".format(
                prev[0] or "—", prev[1] or "—", prev[2] or "—", prev[3] or "—",
                ev["f1"] or "—", ev["k1"] or "—", ev["f2"] or "—", ev["k2"] or "—")
            detail += fmt_changes(prev, ev)
        if ev["f1"] is None and ev["f2"] is None:
            stats["noline"] += 1

        live_tag = " [LIVE]" if ev["live"] else ""
        log_lines.append("{}{} | {}: {}".format(ev["start"], live_tag, ev["match"], detail))

        append_csv(csv_path, [
            ts, ev["id"], ev["match"], ev["start"],
            ev["f1"] or "", ev["k1"] or "",
            ev["f2"] or "", ev["k2"] or "",
            status,
        ])
        # обновляем состояние
        prev_state[ev["id"]] = list(cur)

    append_log(log_path, "\n".join(log_lines))
    save_state(state_path, prev_state)

    summary = ("Проверка {} МСК: матчей {}, новых {}, изменений {}, "
               "без изменений {} (без линии {}).".format(
                   ts, stats["total"], stats["new"], stats["changed"],
                   stats["same"], stats["noline"]))
    if not quiet:
        print(summary)
    return {"ok": True, "stats": stats, "summary": summary}


# ----------------------------------------------------------------------------
# Точка входа
# ----------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Мониторинг фор NBA на Marathonbet: лог изменений по каждой игре.")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--once", action="store_true", help="одна проверка и выход")
    mode.add_argument("--loop", action="store_true",
                      help="вечный цикл: проверка каждые --interval секунд")
    parser.add_argument("--interval", type=int, default=DEFAULT_INTERVAL,
                        help="интервал цикла в секундах (по умолчанию 1800 = 30 мин)")
    parser.add_argument("--url", default=DEFAULT_URL, help="адрес страницы с линией")
    parser.add_argument("--data-dir", default=None,
                        help="папка для history.csv / log.txt / state.json "
                             "(по умолчанию — рядом со скриптом)")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT)
    parser.add_argument("--retries", type=int, default=DEFAULT_RETRIES)
    args = parser.parse_args()

    # Чтобы кириллица не ломалась в консоли Windows
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

    data_dir = args.data_dir or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "marathon_fora_data")

    if args.loop:
        print("Запущен вечный цикл: проверка каждые {} сек. Остановка — Ctrl+C.".format(
            args.interval))
        while True:
            try:
                run_once(data_dir, args.url, timeout=args.timeout, retries=args.retries)
            except KeyboardInterrupt:
                raise
            except Exception as err:  # noqa: BLE001 - цикл не должен падать
                print("Ошибка цикла: {}".format(err), file=sys.stderr)
            try:
                time.sleep(args.interval)
            except KeyboardInterrupt:
                print("Остановлено пользователем.")
                return
    else:
        result = run_once(data_dir, args.url, timeout=args.timeout, retries=args.retries)
        sys.exit(0 if result["ok"] else 1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Остановлено пользователем.")
