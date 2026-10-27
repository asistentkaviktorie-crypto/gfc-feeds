#!/usr/bin/env python3
"""Download the public part of the GFC Google Sheet (Apps Script web app, see apps_script/Code.gs)
and write it as an xlsx with the same tabs/headers that build_feed.py reads.
Never prints cell contents (Actions logs of a public repo are public).
  python3 build/fetch_sheet.py --url "$SHEET_JSON_URL" --out work/GFC.xlsx
  python3 build/fetch_sheet.py --json saved.json --out work/GFC.xlsx      (offline test)"""
import argparse, json, os, sys, time, urllib.request
import openpyxl

NEED_TABS = ["Обєкти", "Тексти", "Переклади", "Параметри", "Сайт", "Фото"]

def get(url):
    last = None
    for attempt in range(4):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "gfc-feeds-action"})
            with urllib.request.urlopen(req, timeout=90) as r:   # follows the script.googleusercontent.com redirect
                return r.read().decode("utf-8")
        except Exception as e:
            last = e; time.sleep(10 * (attempt + 1))
    sys.exit(f"ПОМИЛКА: не вдалося завантажити таблицю ({type(last).__name__}). Перевірте секрет SHEET_JSON_URL і розгортання Apps Script.")

def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--url"); g.add_argument("--json")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    raw = get(a.url) if a.url else open(a.json, encoding="utf-8").read()
    try: doc = json.loads(raw)
    except ValueError: sys.exit("ПОМИЛКА: відповідь Apps Script не JSON (можливо, розгортання має доступ не «Усі» або потрібна повторна авторизація).")
    tabs = doc.get("tabs") or {}
    miss = [t for t in NEED_TABS if t not in tabs]
    if miss: sys.exit(f"ПОМИЛКА: у таблиці немає вкладок {miss}")
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for name in NEED_TABS:
        ws = wb.create_sheet(name)
        for row in tabs[name]:
            ws.append([None if v == "" else v for v in row])
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    wb.save(a.out)
    print("sheet ok:", {n: max(len(tabs[n]) - 1, 0) for n in NEED_TABS})

if __name__ == "__main__":
    main()
