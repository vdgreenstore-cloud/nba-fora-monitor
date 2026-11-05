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
  6. Ведёт архив изменений по игровым дням: data/изменения/ГГГГ-ММ-ДД.txt|.html —
     только игры со сдвигами форы за каждый игровой день; после завершения
     матча в конце его блока появляется итоговый счёт (с пометкой об овертайме)
  7. Ведёт полные отчёты по игровым дням: data/report/Д.М.txt (5.11.txt, 6.11.txt…)
     — все игры дня от появления в линии до закрытия линии (начала матча)

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

# Публичный API ESPN — итоговые счёты завершившихся матчей NBA
ESPN_SCOREBOARD = ("http://site.api.espn.com/apis/site/v2/sports/"
                   "basketball/nba/scoreboard?dates={date}")
# сколько часов после начала матча ждём, прежде чем искать итог (примерно
# длится матч 2-2.5 часа, плюс возможный овертайм и перерывы)
GAME_OVER_AFTER = timedelta(hours=2, minutes=40)
# сколько раз можно пытаться найти счёт матча, прежде чем перестать
SCORE_MAX_TRIES = 8

# русские названия команд у Marathonbet -> аббревиатуры ESPN
TEAM_ABBR = {
    "Атланта Хокс": "ATL",
    "Бостон Селтикс": "BOS",
    "Бруклин Нетс": "BKN",
    "Шарлотт Хорнетс": "CHA",
    "Чикаго Буллз": "CHI",
    "Кливленд Кавальерс": "CLE",
    "Даллас Маверикс": "DAL",
    "Денвер Наггетс": "DEN",
    "Детройт Пистонс": "DET",
    "Голден Стэйт Уорриорз": "GSW",
    "Хьюстон Рокетс": "HOU",
    "Индиана Пэйсерс": "IND",
    "Лос-Анджелес Клипперс": "LAC",
    "Лос-Анджелес Лейкерс": "LAL",
    "Мемфис Гриззлис": "MEM",
    "Майами Хит": "MIA",
    "Милуоки Бакс": "MIL",
    "Миннесота Тимбервулвз": "MIN",
    "Нью-Орлеан Пеликанс": "NOP",
    "Нью-Йорк Никс": "NY",
    "Оклахома-Сити Тандер": "OKC",
    "Орландо Мэджик": "ORL",
    "Филадельфия 76-е": "PHI",
    "Финикс Санз": "PHX",
    "Портленд Трэйл Блэйзерс": "POR",
    "Сакраменто Кингз": "SAC",
    "Сан-Антонио Сперс": "SAS",
    "Торонто Рэпторс": "TOR",
    "Юта Джаз": "UTA",
    "Вашингтон Уизардс": "WAS",
}

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


_HTML_STYLE = (
    "<style>"
    "body{font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif;"
    "margin:14px;color:#1a1a1a;}"
    "h1{font-size:18px;margin:0 0 4px;}"
    ".meta{color:#666;font-size:13px;margin-bottom:8px;line-height:1.5;}"
    ".legend{font-size:12.5px;color:#444;margin:10px 0 16px;}"
    ".legend .sw{display:inline-block;width:12px;height:12px;"
    "background:#fdecea;border:1px solid #d32f2f;vertical-align:-1px;"
    "margin-right:6px;}"
    ".match{margin-bottom:14px;}"
    ".game-head{font-weight:600;margin:14px 0 4px;font-size:14px;}"
    ".line{font-family:Consolas,'Courier New',monospace;font-size:13.5px;"
    "padding:3px 8px;}"
    ".changed{background:#fdecea;color:#b71c1c;font-weight:600;"
    "border-left:3px solid #d32f2f;}"
    ".final{font-family:Consolas,'Courier New',monospace;font-size:13.5px;"
    "padding:3px 8px;color:#1b5e20;font-weight:600;"
    "border-left:3px solid #2e7d32;margin-top:2px;}"
    "</style>"
)


def _write_html(html_path, title, body_lines):
    """Сохраняет HTML-страницу (utf-8) атомарно."""
    out = [
        "<!DOCTYPE html>",
        '<html lang="ru">',
        "<head>",
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        "<title>{}</title>".format(html.escape(title, quote=False)),
        _HTML_STYLE,
        "</head>",
        "<body>",
    ] + list(body_lines) + ["</body>", "</html>"]
    tmp = html_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")
    os.replace(tmp, html_path)


def _load_history_matches(csv_path):
    """Читает history.csv и группирует записи по матчам.

    Возвращает словарь {id: {"name", "start_dated", "start_plain",
    "first_check", "rows": [(время_проверки, строка_csv), ...]}}.
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
    return matches


def _match_blocks(matches):
    """Собирает блоки отчёта по матчам (сортировка по времени начала).

    Возвращает список словарей:
      {"name": str, "start": datetime|None,
       "lines": [(текст_строки, сдвиг_форы: bool)], "shifts": int}
    Сдвигом считается изменение форы относительно ПРЕДЫДУЩЕЙ проверки
    (движение кэфов изменением не считается). Проверки после начала
    матча не включаются.
    """
    def sort_key(mid):
        m = matches[mid]
        start = m["start_dated"] or m["start_plain"]
        return (start or datetime.max, m["name"])

    blocks = []
    for mid in sorted(matches, key=sort_key):
        m = matches[mid]
        start = m["start_dated"] or m["start_plain"]
        rows = sorted(m["rows"], key=lambda x: x[0])
        if start is not None:
            before = [r for r in rows if r[0] <= start]
            if before:  # проверки после начала матча не показываем
                rows = before
        lines = []
        prev_fora = None
        for chk, row in rows:
            f1 = (row.get("фора 1") or "").strip()
            f2 = (row.get("фора 2") or "").strip()
            text = "{} | {}: [{} / {}]".format(
                _fmt_check_dt(chk), m["name"],
                _report_pair(f1, row.get("кэф 1")),
                _report_pair(f2, row.get("кэф 2")))
            changed = prev_fora is not None and (f1, f2) != prev_fora
            lines.append((text, changed))
            prev_fora = (f1, f2)
        blocks.append({
            "mid": mid,
            "name": m["name"],
            "start": start,
            "first": m["first_check"],
            "lines": lines,
            "shifts": sum(1 for _, c in lines if c),
        })
    return blocks


def _write_line_report_html(blocks, html_path, generated_at):
    """HTML-версия общего отчёта: строки со сдвигом форы — красные целиком."""
    upd = (generated_at or datetime.now(MSK)).strftime("%d.%m.%Y %H:%M МСК")
    body = [
        "<h1>Динамика форы NBA — Marathonbet</h1>",
        '<div class="meta">Время проверок московское. Обновлён: {}.<br>'
        "По каждой игре: все проверки от появления в линии до начала матча.<br>"
        "Формат: [фора 1 (кэф 1) / фора 2 (кэф 2)].</div>".format(upd),
        '<div class="legend"><span class="sw"></span>красная строка — фора '
        "изменилась по сравнению с предыдущей проверкой (движение кэфов "
        "изменением не считается)</div>",
    ]
    for block in blocks:
        body.append('<div class="match">')
        for text, changed in block["lines"]:
            cls = "line changed" if changed else "line"
            body.append('<div class="{}">{}</div>'.format(
                cls, html.escape(text, quote=False)))
        body.append("</div>")
    _write_html(html_path, "Динамика форы NBA — Marathonbet", body)


def build_line_report(csv_path, report_path, html_path=None, generated_at=None):
    """Строит текстовый отчёт по истории проверок (history.csv).

    По каждой игре — все её проверки от первого появления в линии
    до начала матча, строка вида:
      05 ноя 03:00 | Вашингтон Уизардс - Атланта Хокс: [+4.5 (1.80) / -4.5 (2.00)]
    Строки, где фора сдвинулась относительно предыдущей проверки, помечаются
    🔴 (в txt) и выделяются красной строкой (в html, если задан html_path).
    Возвращает количество матчей, попавших в отчёт.
    """
    blocks = _match_blocks(_load_history_matches(csv_path))

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
        for text, changed in block["lines"]:
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
    return len(blocks)


def _split_team_names(match_name):
    """'Вашингтон Уизардс - Атланта Хокс' -> ('Вашингтон Уизардс', 'Атланта Хокс').
    Разделитель — ' - ' (с пробелами); названия команд дефис без пробелов
    ('Лос-Анджелес Клипперс') не разрезают."""
    parts = match_name.split(" - ")
    if len(parts) != 2:
        return None
    return parts[0].strip(), parts[1].strip()


def _match_abbrs(match_name):
    """Пара аббревиатур ESPN для матча или None, если команда не распознана."""
    names = _split_team_names(match_name or "")
    if names is None:
        return None
    try:
        return TEAM_ABBR[names[0]], TEAM_ABBR[names[1]]
    except KeyError:
        return None


def _load_results(path):
    """Читает кэш итоговых счётов data/results.json.

    Формат: {id матча: {"score": "126:115"|None, "ot": bool, "tries": int}}
    """
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_results(path, results):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=1, sort_keys=True)
    os.replace(tmp, path)


def _fetch_espn_scoreboard(date, timeout=DEFAULT_TIMEOUT):
    """Итоги NBA за дату (дата по восточно-американскому времени).

    Возвращает словарь {frozenset({abbr1, abbr2}): {"scores": {abbr: очки},
    "ot": bool}} по ЗАВЕРШЁННЫМ матчам этого дня.
    """
    url = ESPN_SCOREBOARD.format(date=date.strftime("%Y%m%d"))
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8", errors="replace"))
    out = {}
    for ev in data.get("events", []):
        try:
            comp = ev["competitions"][0]
            status = ev["status"]
            if not status.get("type", {}).get("completed"):
                continue
            abbrs, scores = set(), {}
            for c in comp["competitors"]:
                abbr = (c.get("team") or {}).get("abbreviation", "").upper()
                abbrs.add(abbr)
                try:
                    scores[abbr] = int(c.get("score"))
                except (TypeError, ValueError):
                    scores[abbr] = None
            if len(abbrs) != 2 or any(v is None for v in scores.values()):
                continue
            ot = status.get("period", 4) >= 5 or "OT" in (
                status.get("type", {}).get("description") or "")
            out[frozenset(abbrs)] = ({"scores": scores, "ot": ot})
        except (KeyError, IndexError, TypeError):
            continue
    return out


def fetch_final_scores(csv_path, results_path, timeout=DEFAULT_TIMEOUT):
    """Дополняет кэш results.json итоговыми счётами завершившихся матчей.

    Матч считается завершённым, если прошло >= 2.5 ч с начала; счёт ищется
    в результатах NBA за соответствующий день (по времени США, с запасным
    предыдущим днём). Счёт записывается в порядке команд Marathonbet:
    "итоговый счёт X:Y" — X очков первой команды в названии матча.
    Возвращает число вновь найденных счётов.
    """
    matches = _load_history_matches(csv_path)
    results = _load_results(results_path)
    now = datetime.now(MSK)

    pending = []  # (id, матч, начало)
    for mid, m in matches.items():
        start = m["start_dated"] or m["start_plain"]
        if start is None or now < start + GAME_OVER_AFTER:
            continue  # ещё не должен был закончиться
        rec = results.get(mid) or {}
        if rec.get("score"):
            continue  # счёт уже известен
        if rec.get("tries", 0) >= SCORE_MAX_TRIES:
            continue  # так и не нашли — больше не ищем
        if _match_abbrs(m["name"]) is None:
            continue  # неизвестное название команды
        pending.append((mid, m["name"], start))
    if not pending:
        return 0

    # у NBA матч начинается 12:00-22:59 по Востоку (ET = МСК - 8 ч);
    # для подстраховки ищем и за предыдущую дату
    dates = set()
    for _, _, start in pending:
        et = start - timedelta(hours=8)
        dates.add(et.date())
        dates.add(et.date() - timedelta(days=1))

    boards = {}
    for d in sorted(dates):
        try:
            boards[d] = _fetch_espn_scoreboard(d, timeout=timeout)
        except Exception as err:  # noqa: BLE001 - сеть/сайт: попробуем в следующий раз
            print("[!] ESPN {}: {}".format(d.strftime("%d.%m.%Y"), err),
                  file=sys.stderr)
            boards[d] = {}

    found = 0
    for mid, name, start in pending:
        abbrs = _match_abbrs(name)
        key = frozenset(abbrs)
        rec = results.setdefault(mid, {"score": None, "ot": False, "tries": 0})
        rec["tries"] = rec.get("tries", 0) + 1
        et = start - timedelta(hours=8)
        hit = None
        for d in (et.date(), et.date() - timedelta(days=1)):
            if key in boards.get(d, {}):
                hit = boards[d][key]
                break
        if hit is not None:
            # счёт в порядке команд из названия матча (первая:вторая)
            sc = hit["scores"]
            rec["score"] = "{}:{}".format(sc[abbrs[0]], sc[abbrs[1]])
            rec["ot"] = bool(hit.get("ot"))
            found += 1
    _save_results(results_path, results)
    return found


def build_changes_archive(csv_path, out_dir, generated_at=None, results_path=None):
    """Архив изменений по игровым дням.

    Для каждого игрового дня (дата начала матчей по МСК), в котором хотя бы
    у одной игры сдвинулась фора, создаёт в папке out_dir пару файлов:
      ГГГГ-ММ-ДД.txt и ГГГГ-ММ-ДД.html
    Внутри — только игры со сдвигами, все их проверки от появления в линии
    до начала матча; строки со сдвигом помечены 🔴 / красным; после завершения
    матча в конце его блока добавляется строка «Итоговый счёт: X:Y»
    (с пометкой об овертайме, если он был) из кэша results.json.
    Возвращает список созданных файлов.
    """
    blocks = _match_blocks(_load_history_matches(csv_path))
    results = _load_results(results_path)

    days = {}  # дата начала матчей -> список блоков
    for block in blocks:
        if block["shifts"] <= 0 or block["start"] is None:
            continue
        days.setdefault(block["start"].date(), []).append(block)

    os.makedirs(out_dir, exist_ok=True)
    written = []
    for day in sorted(days):
        day_blocks = days[day]
        stem = day.strftime("%Y-%m-%d")
        txt_path = os.path.join(out_dir, stem + ".txt")
        html_path = os.path.join(out_dir, stem + ".html")
        _write_day_txt(day, day_blocks, txt_path, results)
        _write_day_html(day, day_blocks, html_path, results)
        written.extend([txt_path, html_path])
    return written


def _day_head_lines(day, day_blocks):
    """Шапка дневного файла архива (для txt)."""
    total_shifts = sum(b["shifts"] for b in day_blocks)
    return [
        "ИГРОВОЙ ДЕНЬ {} (МСК) — игры со сдвигами форы".format(
            day.strftime("%d.%m.%Y")),
        "Мониторинг линии: от появления игры до её начала; "
        "движение кэфов изменением не считается.",
        "Игр со сдвигами: {} | Сдвигов форы: {}".format(
            len(day_blocks), total_shifts),
    ]


def _final_lines(block, results):
    """Строка (или пустой список) с итоговым счётом после блока матча.

    Счёт известен, если матч завершился и итог найден в кэше results.json
    (results передаются как {id матча: {...}}).
    """
    rec = (results or {}).get(block.get("mid")) or {}
    score = rec.get("score")
    if not score:
        return []
    line = "Итоговый счёт: {}".format(score)
    if rec.get("ot"):
        line += " (овертайм)"
    return [line]


def _write_day_txt(day, day_blocks, txt_path, results=None):
    lines = _day_head_lines(day, day_blocks)
    lines.append("")
    for block in day_blocks:
        start_txt = (_fmt_check_dt(block["start"].replace(second=0, microsecond=0))
                     if block["start"] else "—")
        lines.append("=== {} | начало {} | сдвигов форы: {} ===".format(
            block["name"], start_txt, block["shifts"]))
        for text, changed in block["lines"]:
            lines.append(("🔴 " + text) if changed else text)
        lines.extend(_final_lines(block, results))
        lines.append("")
    while lines and lines[-1] == "":
        lines.pop()
    tmp = txt_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    os.replace(tmp, txt_path)


def _day_of(block):
    """Игровой день блока: дата начала матча (или первой проверки, если
    время начала в истории так и не появилось)."""
    return block["start"].date() if block["start"] else block["first"].date()


def _day_file_name(day):
    """Имя дневного файла отчёта: 5.11.txt, 6.11.txt (без ведущих нулей)."""
    return "{}.{}.txt".format(day.day, day.month)


def build_day_reports(csv_path, out_dir, generated_at=None):
    """Полные отчёты по игровым дням: out_dir/Д.М.txt (5.11.txt, 6.11.txt…).

    За каждый игровой день (дата начала матчей по МСК) создаёт файл,
    где собраны ВСЕ игры этого дня и все их проверки от первого появления
    в линии до закрытия линии (последней проверки перед началом матча).
    Проверки после начала матча не включаются; строки со сдвигом форы
    помечены 🔴. Пересобирается целиком после каждой проверки.
    Возвращает список созданных файлов.
    """
    blocks = _match_blocks(_load_history_matches(csv_path))

    days = {}  # дата игрового дня -> список блоков (отсортированы по началу)
    for block in blocks:
        days.setdefault(_day_of(block), []).append(block)

    os.makedirs(out_dir, exist_ok=True)
    wanted = set()
    written = []
    for day in sorted(days):
        path = os.path.join(out_dir, _day_file_name(day))
        _write_gameday_txt(day, days[day], path, generated_at)
        written.append(path)
        wanted.add(os.path.basename(path))
    # дневные файлы прошлых запусков, которых больше нет в истории,
    # удаляем — чтобы в папке остались только актуальные дни
    for fn in os.listdir(out_dir):
        if fn.endswith(".txt") and fn not in wanted:
            os.remove(os.path.join(out_dir, fn))
    return written


def _write_gameday_txt(day, day_blocks, txt_path, generated_at=None):
    lines = [
        "ИГРОВОЙ ДЕНЬ {} (МСК) — все игры дня".format(day.strftime("%d.%m.%Y")),
        "По каждой игре: все проверки от появления в линии до закрытия "
        "(последняя проверка перед началом матча).",
        "Формат линии: [фора 1 (кэф 1) / фора 2 (кэф 2)].",
        "🔴 — фора изменилась относительно предыдущей проверки "
        "(движение кэфов изменением не считается).",
        "Обновлён: {}".format(
            (generated_at or datetime.now(MSK)).strftime("%d.%m.%Y %H:%M МСК")),
        "Игр: {} | из них со сдвигами форы: {}".format(
            len(day_blocks), sum(1 for b in day_blocks if b["shifts"])),
        "",
    ]
    for block in day_blocks:
        start_txt = (_fmt_check_dt(block["start"].replace(second=0, microsecond=0))
                     if block["start"] else "—")
        lines.append("=== {} | начало {} | сдвигов форы: {} ===".format(
            block["name"], start_txt, block["shifts"]))
        for text, changed in block["lines"]:
            lines.append(("🔴 " + text) if changed else text)
        lines.append("")
    while lines and lines[-1] == "":
        lines.pop()
    tmp = txt_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    os.replace(tmp, txt_path)


def _write_day_html(day, day_blocks, html_path, results=None):
    body = [
        "<h1>Изменения форы NBA — игровой день {}</h1>".format(
            day.strftime("%d.%m.%Y")),
        '<div class="meta">Время московское. Здесь только игры, у которых '
        "сдвинулась фора: все их проверки от появления в линии до начала "
        "матча.</div>",
        '<div class="legend"><span class="sw"></span>красная строка — фора '
        "изменилась по сравнению с предыдущей проверкой (движение кэфов "
        "изменением не считается)</div>",
    ]
    for block in day_blocks:
        start_txt = (_fmt_check_dt(block["start"].replace(second=0, microsecond=0))
                     if block["start"] else "—")
        body.append(
            '<div class="game-head">{} · начало {} · сдвигов форы: {}</div>'
            .format(html.escape(block["name"], quote=False),
                    start_txt, block["shifts"]))
        for text, changed in block["lines"]:
            cls = "line changed" if changed else "line"
            body.append('<div class="{}">{}</div>'.format(
                cls, html.escape(text, quote=False)))
        for text in _final_lines(block, results):
            body.append('<div class="final">{}</div>'.format(
                html.escape(text, quote=False)))
    _write_html(html_path,
                "Изменения форы NBA — {}".format(day.strftime("%d.%m.%Y")),
                body)


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

    # итоговые счёты завершившихся матчей (кэш data/results.json)
    results_path = os.path.join(data_dir, "results.json")
    try:
        found_scores = fetch_final_scores(csv_path, results_path)
        if found_scores:
            print("Найдено итоговых счётов: {}".format(found_scores))
    except Exception as err:  # noqa: BLE001 - счёты не должны ломать проверку
        print("[!] не удалось получить итоговые счёты: {}".format(err),
              file=sys.stderr)

    # архив изменений по игровым дням: data/изменения/ГГГГ-ММ-ДД.txt|.html
    try:
        build_changes_archive(csv_path, os.path.join(data_dir, "изменения"),
                              results_path=results_path)
    except Exception as err:  # noqa: BLE001 - архив не должен ломать проверку
        print("[!] не удалось построить архив изменений: {}".format(err),
              file=sys.stderr)

    # полные отчёты по игровым дням: data/report/5.11.txt, 6.11.txt, ...
    try:
        build_day_reports(csv_path, os.path.join(data_dir, "report"))
    except Exception as err:  # noqa: BLE001 - отчёты не должны ломать проверку
        print("[!] не удалось построить отчёты по игровым дням: {}".format(err),
              file=sys.stderr)

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
