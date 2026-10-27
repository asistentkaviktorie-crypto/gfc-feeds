#!/usr/bin/env python3
"""TEST ONLY: produce the same JSON the Apps Script web app (apps_script/Code.gs) returns, from a local xlsx."""
import json, sys, openpyxl
OBJ = ['Проєкт', 'Ціна €', 'Площа заг., м²', 'Житлова, м²', 'Ділянка, м²', 'Спальні', 'Ванні',
       'Realting ID', 'Realting статус', 'У фід Realting', 'У фід JamesEdition', 'У фід Prian']
PUB = ['Тексти', 'Переклади', 'Параметри', 'Сайт', 'Фото']
YES = ['так', 'yes', 'да', '1', 'true']
wb = openpyxl.load_workbook(sys.argv[1], data_only=False)
def vals(ws): return [["" if c is None else c for c in r] for r in ws.iter_rows(values_only=True) if any(c not in (None, "") for c in r)]
v = vals(wb.worksheets[0]); hdr = [str(h).strip() for h in v[0]]
known = {str(r[0]).strip().upper() for n in ("Параметри", "Сайт") for r in vals(wb[n])[1:] if r[0] != ""}
rid = hdr.index('Realting ID'); fl = [hdr.index(c) for c in ('У фід Realting', 'У фід JamesEdition', 'У фід Prian')]
rf = hdr.index('У фід Realting')
rows = [OBJ]
for r in v[1:]:
    r = list(r) + [""] * (len(hdr) - len(r))
    # formulas like "=P10" are evaluated by Google; emulate "same as Realting"
    for i in fl:
        if str(r[i]).startswith("="): r[i] = r[rf]
    i = str(r[rid]).strip()
    if not i: continue
    if not any(str(r[j]).strip().lower() in YES for j in fl) and i.upper() not in known: continue
    rows.append([r[hdr.index(c)] if c in hdr else "" for c in OBJ])
sent = {str(r[7]).strip().upper() for r in rows[1:]}
tabs = {n: vals(wb[n]) for n in PUB}
tabs["Тексти"] = [tabs["Тексти"][0]] + [r for r in tabs["Тексти"][1:] if str(r[0]).strip().upper() in sent]
out = {"version": 1, "tabs": {"Обєкти": rows, **tabs}}
json.dump(out, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, default=str)
print("rows Обєкти:", len(rows) - 1)
