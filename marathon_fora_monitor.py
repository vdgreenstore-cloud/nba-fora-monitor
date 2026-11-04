#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Мониторинг фор (гандикапов) на матчи NBA у букмекера Marathonbet.

Что делает:
  1. Скачивает страницу https://www.marathonbet.ru/su/betting/Basketball/NBA
  2. По каждой игре извлекает главную фору и коэффициенты на неё
  3. Сравнивает с предыдущим запуском и построчно записывает по каждой игре,
     изменилась линия или нет. Изменением считается только сдвиг ФОРЫ;
     коэффициенты записываются в историю, но их движение не считается
     изменением:
        - в читаемый лог   log.txt
        - в историю        history.csv (открывается в Excel)
  4. Текущий снимок хранит в state.json
  5. Строит отчёты report.txt (🔴 = сдвиг форы) и report.html (сдвиг форы
     выделен красной строкой) — по каждой игре от появления до начала матча

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
# Текстовый отчёт: динамика форы по каждой игре (report.txt)
# ----------------------------------------------------------------------------

MONTHS_RU = ["янв", "фев", "мар", "апр", "мая", "июн",
             "июл", "авг", "сен", "окт", "ноя", "дек"]


def _parse_check_dt(value):
    """'04.11.2026 23:57:35' -> datetime."""
    return datetime.strptime(value.strip(), "%d.%m.%Y %H:%M:%S")


def _fmt_check_dt(dt):
    """datetime -> '05 ноя 03:00'."""
    return "{:02d} {} {:02d}:{:02d}".format(
        dt.day, MONTHS_RU[dt.month - 1], dt.hour, dt.minute)


def _normalize_start(raw, check_dt):
    """Приводит время начала матча к datetime.

    На сайте два варианта: '05 ноя 03:00' (с датой) и просто '03:00'
    (когда матч в день проверки). Во втором случае дата берётся от даты
    проверки; если такое время раньше самой проверки — матч на следующий день.
    """
    raw = (raw or "").strip()
    if not raw:
        return None
    m = re.match(r"^(\d{1,2})\s+([а-яё]+)\.?\s+(\d{1,2}):(\d{2})$",
                 raw, re.IGNORECASE)
    if m:
        mon_name = m.group(2).lower()[:3]
        if mon_name not in MONTHS_RU:
            return None
        month = MONTHS_RU.index(mon_name) + 1
        year = check_dt.year
        if month == 1 and check_dt.month == 12:  # переход через новый год
            year += 1
        try:
            return datetime(year, month, int(m.group(1)),
                            int(m.group(3)), int(m.group(4)))
        except ValueError:
            return None
    m = re.match(r"^(\d{1,2}):(\d{2})$", raw)
    if m:
        start = check_dt.replace(hour=int(m.group(1)), minute=int(m.group(2)),
                                 second=0, microsecond=0)
        if start < check_dt - timedelta(minutes=5):
            start += timedelta(days=1)
        return start
    return None


def _report_pair(fora, price):
    """'фора (кэф)' с аккуратной заменой пустых значений на '—'."""
    fora = (fora or "").strip()
    price = (price or "").strip()
    if not fora and not price:
        return "—"
    return "{} ({})".format(fora or "—", price or "—")


def _write_line_report_html(blocks, html_path, generated_at):
    """Пишет HTML-версию отчёта: строки со сдвигом форы выделены красным
    целиком (открывается в любом браузере)."""
    upd = (generated_at or datetime.now(MSK)).strftime("%d.%m.%Y %H:%M МСК")
    out = [
        "<!DOCTYPE html>",
        '<html lang="ru">',
        "<head>",
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        "<title>Динамика форы NBA — Marathonbet</title>",
        "<style>",
        "body{font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif;"
        "margin:14px;color:#1a1a1a;}",
        "h1{font-size:18px;margin:0 0 4px;}",
        ".meta{color:#666;font-size:13px;margin-bottom:8px;line-height:1.5;}",
        ".legend{font-size:12.5px;color:#444;margin:10px 0 16px;}",
        ".legend .sw{display:inline-block;width:12px;height:12px;"
        "background:#fdecea;border:1px solid #d32f2f;vertical-align:-1px;"
        "margin-right:6px;}",
        ".match{margin-bottom:14px;}",
        ".line{font-family:Consolas,'Courier New',monospace;font-size:13.5px;"
        "padding:3px 8px;}",
        ".changed{background:#fdecea;color:#b71c1c;font-weight:600;"
        "border-left:3px solid #d32f2f;}",
        "</style>",
        "</head>",
        "<body>",
        "<h1>Динамика форы NBA — Marathonbet</h1>",
        '<div class="meta">Время проверок московское. Обновлён: {}.<br>'
        "По каждой игре: все проверки от появления в линии до начала матча.<br>"
        "Формат: [фора 1 (кэф 1) / фора 2 (кэф 2)].</div>".format(upd),
        '<div class="legend"><span class="sw"></span>красная строка — фора '
        "изменилась по сравнению с предыдущей проверкой (движение кэфов "
        "изменением не считается)</div>",
    ]
    for block in blocks:
        out.append('<div class="match">')
        for text, changed in block:
            cls = "line changed" if changed else "line"
            out.append('<div class="{}">{}</div>'.format(
                cls, html.escape(text, quote=False)))
        out.append("</div>")
    out.extend(["</body>", "</html>"])
    tmp = html_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")
    os.replace(tmp, html_path)


def build_line_report(csv_path, report_path, html_path=None, generated_at=None):
    """Строит текстовый отчёт по истории проверок (history.csv).

    По каждой игре — все её проверки от первого появления в линии
    до начала матча, строка вида:
      05 ноя 03:00 | Вашингтон Уизардс - Атланта Хокс: [+4.5 (1.80) / -4.5 (2.00)]
    Строки, где фора сдвинулась относительно предыдущей проверки, помечаются
    🔴 (в txt) и выделяются красной строкой (в html, если задан html_path).
    Возвращает количество матчей, попавших в отчёт.
    """
    matches = {}
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            try:
                chk = _parse_check_dt(row.get("время проверки") or "")
            except ValueError:
                continue
            mid = (row.get("id матча") or row.get("матч") or "").strip()
            m = matches.setdefault(mid, {
                "name": (row.get("матч") or "").strip(),
                "start_dated": None,   # начало из записи с датой '05 ноя 03:00'
                "start_plain": None,   # начало из записи только со временем
                "first_check": chk,
                "rows": [],
            })
            m["rows"].append((chk, row))
            if chk < m["first_check"]:
                m["first_check"] = chk

            raw_start = (row.get("начало (МСК)") or "").strip()
            norm = _normalize_start(raw_start, chk)
            if norm is None:
                continue
            if re.search(r"[а-яё]", raw_start, re.IGNORECASE):
                if m["start_dated"] is None or norm < m["start_dated"]:
                    m["start_dated"] = norm
            elif m["start_plain"] is None or norm < m["start_plain"]:
                m["start_plain"] = norm

    def sort_key(mid):
        m = matches[mid]
        start = m["start_dated"] or m["start_plain"]
        return (start or datetime.max, m["name"])

    blocks = []  # блоки: список пар (строка, изменилась_ли_фора)
    for mid in sorted(matches, key=sort_key):
        m = matches[mid]
        start = m["start_dated"] or m["start_plain"]
        rows = sorted(m["rows"], key=lambda x: x[0])
        if start is not None:
            before = [r for r in rows if r[0] <= start]
            if before:  # проверки после начала матча не показываем
                rows = before
        block = []
        prev_fora = None
        for chk, row in rows:
            f1 = (row.get("фора 1") or "").strip()
            f2 = (row.get("фора 2") or "").strip()
            text = "{} | {}: [{} / {}]".format(
                _fmt_check_dt(chk), m["name"],
                _report_pair(f1, row.get("кэф 1")),
                _report_pair(f2, row.get("кэф 2")))
            # сдвиг ФОРЫ относительно предыдущей проверки; кэфы не считаются
            changed = prev_fora is not None and (f1, f2) != prev_fora
            block.append((text, changed))
            prev_fora = (f1, f2)
        blocks.append(block)

    lines = [
        "Динамика форы NBA — Marathonbet (время проверок московское)",
        "По каждой игре: все проверки от появления в линии до начала матча.",
        "Формат линии: [фора 1 (кэф 1) / фора 2 (кэф 2)].",
        "🔴 — фора изменилась относительно предыдущей проверки "
        "(движение кэфов изменением не считается).",
        "Обновлён: {}".format(
            (generated_at or datetime.now(MSK)).strftime("%d.%m.%Y %H:%M МСК")),
        "",
    ]
    for block in blocks:
        for text, changed in block:
            lines.append(("🔴 " + text) if changed else text)
        lines.append("")

    while lines and lines[-1] == "":
        lines.pop()

    tmp = report_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    os.replace(tmp, report_path)

    if html_path:
        _write_line_report_html(blocks, html_path, generated_at)
    return len(matches)


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
    """Перечисляет сдвиги ФОРЫ по сравнению с прошлым снимком (кэфы не считаем)."""
    labels = [
        ("фора 1", 0),
        ("фора 2", 2),
    ]
    diffs = []
    for label, idx in labels:
        cur_val = snapshot_tuple(ev)[idx]
        old_val = prev[idx]
        if cur_val != old_val:
            diffs.append("{}: {} → {}".format(label, old_val if old_val else "—",
                                              cur_val if cur_val else "—"))
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
        cur = snapshot_tuple(ev)  # полное состояние: [фора1, кэф1, фора2, кэф2]
        prev = prev_state.get(ev["id"])
        # сравниваем ТОЛЬКО форы (индексы 0 и 2); кэфы не считаются изменением
        same_fora = (isinstance(prev, list) and len(prev) >= 3
                     and (ev["f1"], ev["f2"]) == (prev[0], prev[2]))

        if prev is None:
            status = "НОВЫЙ МАТЧ"
            stats["new"] += 1
            detail = fmt_fora(ev)
        elif same_fora:
            status = "БЕЗ ИЗМЕНЕНИЙ"
            stats["same"] += 1
            detail = fmt_fora(ev)
        else:
            status = "ИЗМЕНЕНИЕ"
            stats["changed"] += 1
            detail = "было [{}/{}] → стало [{}/{}]; ".format(
                prev[0] or "—", prev[2] or "—",
                ev["f1"] or "—", ev["f2"] or "—")
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

    # отчёты: динамика форы по каждой игре от появления до начала (txt + html)
    try:
        build_line_report(csv_path, os.path.join(data_dir, "report.txt"),
                          html_path=os.path.join(data_dir, "report.html"))
    except Exception as err:  # noqa: BLE001 - отчёт не должен ломать проверку
        print("[!] не удалось построить отчёты: {}".format(err), file=sys.stderr)

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
