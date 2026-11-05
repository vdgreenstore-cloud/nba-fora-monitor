#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Проверка build_line_report / build_changes_archive / build_day_reports:
формат строки, фильтр проверок после начала матча, пометка 🔴/красной строкой
ТОЛЬКО сдвигов форы (кэфы не считаются), дневные файлы 5.11.txt, 6.11.txt,
итоговые счёты, архив только .txt, точное расписание (шлюз получасовых слотов)."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from marathon_fora_monitor import (build_line_report, build_changes_archive,  # noqa: E402
                                   build_day_reports, _final_lines,
                                   _match_abbrs, _save_results)

HEADER = "время проверки;id матча;матч;начало (МСК);фора 1;кэф 1;фора 2;кэф 2;статус"
ROWS = [
    # A: старт '05 ноя 03:00'; вторая проверка — сдвиг форы (должна стать красной)
    "04.11.2026 23:57:35;1;А - Б;05 ноя 03:00;+4.5;1.80;-4.5;2.00;НОВЫЙ МАТЧ",
    "05.11.2026 01:00:00;1;А - Б;03:00;+3.5;1.80;-3.5;2.00;ИЗМЕНЕНИЕ",
    "05.11.2026 05:00:00;1;А - Б;03:00;+3.5;1.80;-3.5;2.00;БЕЗ ИЗМЕНЕНИЙ",
    # B: только время старта '06:30' (старт 05.11 06:30); одна проверка в отчёте
    "04.11.2026 23:57:35;2;В - Г;06:30;+1.5;1.90;-1.5;1.90;НОВЫЙ МАТЧ",
    "05.11.2026 06:45:00;2;В - Г;06:30;+1.5;1.90;-1.5;1.90;БЕЗ ИЗМЕНЕНИЙ",
    # C: без форы
    "04.11.2026 23:57:35;3;Д - Е;05 ноя 05:00;;;;;НОВЫЙ МАТЧ",
    # D: кэфы сдвинулись, форы те же — НЕ красная
    "04.11.2026 23:57:35;4;Ж - З;05 ноя 05:30;+2.5;1.90;-2.5;1.90;НОВЫЙ МАТЧ",
    "05.11.2026 01:30:00;4;Ж - З;05:30;+2.5;1.95;-2.5;1.85;БЕЗ ИЗМЕНЕНИЙ",
    # E: матч следующего игрового дня (6.11)
    "04.11.2026 23:57:35;5;И - К;06 ноя 03:00;+2.0;1.85;-2.0;1.95;НОВЫЙ МАТЧ",
    "05.11.2026 01:00:00;5;И - К;03:00;+2.0;1.85;-2.0;1.95;БЕЗ ИЗМЕНЕНИЙ",
]


def main():
    tmpdir = tempfile.mkdtemp()
    csv_path = os.path.join(tmpdir, "history.csv")
    with open(csv_path, "w", encoding="utf-8-sig") as f:
        f.write(HEADER + "\n")
        for r in ROWS:
            f.write(r + "\n")
    out = os.path.join(tmpdir, "report.txt")
    html_out = os.path.join(tmpdir, "report.html")
    n = build_line_report(csv_path, out, html_path=html_out)
    text = open(out, encoding="utf-8").read()
    html_text = open(html_out, encoding="utf-8").read()
    print(text)
    print("--- HTML (фрагмент) ---")
    print("\n".join(html_text.splitlines()[:5]))
    ok = True

    def fail(msg):
        nonlocal_ok[0] = False
        print("FAIL: " + msg)

    nonlocal_ok = [True]

    if n != 5:
        fail("матчей {} != 5".format(n))
    if "А – Б" not in text:
        fail("нет шапки матча A")
    if "04 ноя 23:57   [+4.5 (1.80) / -4.5 (2.00)]" not in text:
        fail("строка матча A не найдена или неверный формат")
    if not any(l.startswith("🔴  05 ноя 01:00")
               and "[+3.5 (1.80) / -3.5 (2.00)]" in l
               for l in text.splitlines()):
        fail("сдвиг форы не помечен 🔴 (A, вторая проверка)")
    if any(l.startswith("🔴  04 ноя 23:57") for l in text.splitlines()):
        fail("первая проверка помечена как изменение (A)")
    if any(l.startswith("🔴  05 ноя 01:30") for l in text.splitlines()):
        fail("движение кэфов помечено как изменение (D)")
    if "05 ноя 01:30   [+2.5 (1.95) / -2.5 (1.85)]" not in text:
        fail("строка с новыми кэфами потеряна (D)")
    if "05 ноя 05:00   [" in text:
        fail("проверка после начала матча попала в отчёт (A)")
    if "05 ноя 06:45   [" in text:
        fail("проверка после начала матча попала в отчёт (B)")
    if "[— / —]" not in text:
        fail("матч без форы отображён неверно")
    if not (text.find("А – Б") < text.find("Д – Е") < text.find("Ж – З")
            < text.find("В – Г")):
        fail("неверный порядок матчей (А 03:00, Д 05:00, Ж 05:30, В 06:30)")
    # HTML-версия
    changed_cnt = html_text.count('class="line changed"')
    if changed_cnt != 1:
        fail("в HTML красных строк {} != 1".format(changed_cnt))
    if "🔴" in html_text:
        fail("в HTML не должно быть маркера 🔴 (он только для txt)")
    if "05 ноя 01:00   [+3.5 (1.80) / -3.5 (2.00)]" not in html_text:
        fail("в HTML нет строки сдвига форы")
    if "А – Б" not in html_text:
        fail("в HTML нет шапки матча A")
    # архив изменений по игровым дням
    arch_dir = os.path.join(tmpdir, "изменения")
    # кэш итоговых счётов: матч 1 (А - Б) завершился 126:115,
    # матч 4 (Ж - З, без сдвигов — в архив не попадает) — 100:99 (OT)
    results_path = os.path.join(tmpdir, "results.json")
    _save_results(results_path, {
        "1": {"score": "126:115", "ot": False, "tries": 1},
        "4": {"score": "100:99", "ot": True, "tries": 1},
    })
    written = build_changes_archive(csv_path, arch_dir, results_path=results_path)
    names = sorted(os.path.basename(p) for p in written)
    if names != ["2026-11-05.txt"]:
        fail("файлы архива: {} != ['2026-11-05.txt']".format(names))
    day_txt = open(os.path.join(arch_dir, "2026-11-05.txt"), encoding="utf-8").read()
    if "ИГРОВОЙ ДЕНЬ 05.11.2026 (МСК) — ИГРЫ СО СДВИГАМИ ФОРЫ" not in day_txt:
        fail("неверная шапка дневного файла")
    if "Игр со сдвигами: 1 · сдвигов форы: 1" not in day_txt:
        fail("неверная статистика в шапке архива")
    if "1. А – Б" not in day_txt:
        fail("в архиве нет заголовка матча A")
    if "начало 05 ноя 03:00 · проверок: 2 · сдвигов форы: 1" not in day_txt:
        fail("неверная строка статистики матча A")
    if not any(l.startswith("🔴  05 ноя 01:00")
               and "[+3.5 (1.80) / -3.5 (2.00)]" in l
               for l in day_txt.splitlines()):
        fail("в архиве нет 🔴 строки сдвига")
    if "Ж - З" in day_txt:
        fail("в архив попала игра без сдвига форы (D)")
    if os.path.exists(os.path.join(arch_dir, "2026-11-05.html")):
        fail("архив создал .html, хотя должен только .txt")
    # итоговый счёт после матча
    if "Итоговый счёт: 126:115" not in day_txt:
        fail("в архиве нет итогового счёта (А - Б)")
    if "Итоговый счёт: 126:115 (овертайм)" in day_txt:
        fail("овертайм отмечен без овертайма (А - Б)")
    if "Итоговый счёт: 100:99" in day_txt:
        fail("счёт игры без сдвигов попал в архив (Ж - З)")
    fl = _final_lines({"mid": "5"}, {"5": {"score": "100:99", "ot": True}})
    if fl != ["✅  Итоговый счёт: 100:99 (овертайм)"]:
        fail("неверная строка счёта с овертаймом: {}".format(fl))
    if _final_lines({"mid": "9"}, {"5": {"score": "100:99", "ot": True}}) != []:
        fail("для матча без счёта вернулась непустая строка")
    # распознавание команд
    if _match_abbrs("Вашингтон Уизардс - Атланта Хокс") != ("WAS", "ATL"):
        fail("неверные аббревиатуры (Вашингтон/Атланта)")
    if _match_abbrs("Лос-Анджелес Клипперс - Майами Хит") != ("LAC", "MIA"):
        fail("дефис без пробелов разрезал название (Клипперс)")
    if _match_abbrs("Неизвестная команда - Майами Хит") is not None:
        fail("неизвестная команда должна давать None")
    # парсер ESPN (фейковый ответ: 1 завершён с OT, 1 незавершён)
    import json as _json
    import marathon_fora_monitor as mfm
    fake = _json.dumps({"events": [
        {"competitions": [{"competitors": [
            {"homeAway": "home", "team": {"abbreviation": "WAS"}, "score": "105"},
            {"homeAway": "away", "team": {"abbreviation": "ATL"}, "score": "120"},
        ]}],
         "status": {"period": 5, "type": {"completed": True,
                                          "description": "Final/OT"}}},
        {"competitions": [{"competitors": [
            {"homeAway": "home", "team": {"abbreviation": "BOS"}, "score": "10"},
            {"homeAway": "away", "team": {"abbreviation": "MIL"}, "score": "8"},
        ]}],
         "status": {"period": 2, "type": {"completed": False,
                                          "description": "2nd Quarter"}}},
    ]}).encode("utf-8")

    class _Resp:
        def __enter__(self):
            return self
        def __exit__(self, *a):
            return False
        def read(self):
            return fake
    orig_urlopen = mfm.urllib.request.urlopen
    mfm.urllib.request.urlopen = lambda req, timeout=None: _Resp()
    try:
        from datetime import date as _date
        board = mfm._fetch_espn_scoreboard(_date(2026, 11, 4))
    finally:
        mfm.urllib.request.urlopen = orig_urlopen
    key = frozenset(("WAS", "ATL"))
    if key not in board:
        fail("ESPN-парсер: завершённый матч не найден")
    else:
        if board[key]["scores"] != {"WAS": 105, "ATL": 120}:
            fail("ESPN-парсер: неверные счёты: {}".format(board[key]["scores"]))
        if not board[key]["ot"]:
            fail("ESPN-парсер: овертайм не определён (period 5)")
    if frozenset(("BOS", "MIL")) in board:
        fail("ESPN-парсер: незавершённый матч попал в итоги")
    # точное расписание: одна проверка на получасовой слот
    from datetime import datetime as _dt
    if mfm.slot_start(_dt(2026, 11, 5, 10, 14, 59)) != _dt(2026, 11, 5, 10, 0):
        fail("slot_start: 10:14:59 должен давать 10:00")
    if mfm.slot_start(_dt(2026, 11, 5, 10, 47, 0)) != _dt(2026, 11, 5, 10, 30):
        fail("slot_start: 10:47 должен давать 10:30")
    if mfm.slot_start(_dt(2026, 11, 5, 23, 59, 59)) != _dt(2026, 11, 5, 23, 30):
        fail("slot_start: 23:59 должен давать 23:30")
    state_path = os.path.join(tmpdir, "state.json")
    if not mfm.should_check_now(state_path, now=_dt(2026, 11, 5, 10, 14)):
        fail("пустой state: проверка должна быть разрешена")
    mfm.save_state(state_path, {"1": ["+1.5", "1.9", "-1.5", "1.9"],
                                "last_check": "2026-11-05 10:05:00"})
    if mfm.should_check_now(state_path, now=_dt(2026, 11, 5, 10, 14)):
        fail("тот же слот (10:05 -> 10:14): проверка должна пропускаться")
    if mfm.should_check_now(state_path, now=_dt(2026, 11, 5, 10, 29, 59)):
        fail("тот же слот (10:05 -> 10:29:59): проверка должна пропускаться")
    if not mfm.should_check_now(state_path, now=_dt(2026, 11, 5, 10, 30)):
        fail("новый слот (10:05 -> 10:30): проверка должна выполняться")
    if not mfm.should_check_now(state_path, now=_dt(2026, 11, 5, 11, 0, 1)):
        fail("новый слот (10:05 -> 11:00): проверка должна выполняться")
    mfm.save_state(state_path, {"last_check": "мусор"})
    if not mfm.should_check_now(state_path, now=_dt(2026, 11, 5, 10, 14)):
        fail("испорченная метка: проверка должна выполняться")
    # полные дневные отчёты data/report/5.11.txt, 6.11.txt
    rep_dir = os.path.join(tmpdir, "report")
    written_rep = build_day_reports(csv_path, rep_dir)
    rep_names = sorted(os.path.basename(p) for p in written_rep)
    if rep_names != ["5.11.txt", "6.11.txt"]:
        fail("файлы дневных отчётов: {} != ['5.11.txt', '6.11.txt']".format(rep_names))
    d5 = open(os.path.join(rep_dir, "5.11.txt"), encoding="utf-8").read()
    d6 = open(os.path.join(rep_dir, "6.11.txt"), encoding="utf-8").read()
    if "ИГРОВОЙ ДЕНЬ 05.11.2026 (МСК) — ВСЕ ИГРЫ ДНЯ" not in d5:
        fail("неверная шапка дневного отчёта 5.11.txt")
    if "Игр: 4 · со сдвигами форы: 1" not in d5:
        fail("неверная статистика в шапке 5.11.txt")
    for g in ("А – Б", "В – Г", "Д – Е", "Ж – З"):
        if g not in d5:
            fail("в 5.11.txt нет игры {}".format(g))
    if "И – К" in d5:
        fail("в 5.11.txt попала игра другого дня (И - К)")
    if not any(l.startswith("🔴  05 ноя 01:00")
               and "[+3.5 (1.80) / -3.5 (2.00)]" in l
               for l in d5.splitlines()):
        fail("в 5.11.txt сдвиг форы не помечен 🔴")
    if "05 ноя 05:00   [" in d5:
        fail("в 5.11.txt попала проверка после начала матча (А)")
    if "05 ноя 06:45   [" in d5:
        fail("в 5.11.txt попала проверка после закрытия линии (В - Г)")
    if "ИГРОВОЙ ДЕНЬ 06.11.2026" not in d6 or "И – К" not in d6:
        fail("6.11.txt должен содержать только И - К")
    if "А – Б" in d6:
        fail("в 6.11.txt попала игра другого дня (А - Б)")
    ok = nonlocal_ok[0]
    print("OK" if ok else "ЕСТЬ ОШИБКИ")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
