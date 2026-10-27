#!/usr/bin/env python3
"""Safety check before publishing. Compares the new build (--new) with what is live now (--old = repo root).
Refuses (exit 1) when:
  * a feed has 0 listings, or the number of listings drops by more than 2 or more than 25 % (unless --allow-drop так);
  * any listing has an empty/zero price, no photos, or a photo file that does not exist;
  * gfc_realting.xml fails the Kyero v3 XSD.
Warnings (do not block): price in the sheet differs from the price written in the text, builder warnings.
Writes a Ukrainian summary to $GITHUB_STEP_SUMMARY when available."""
import argparse, json, os, re, sys
import xml.etree.ElementTree as ET
HERE = os.path.dirname(os.path.abspath(__file__))

def counts(d):
    c = {}
    p = os.path.join(d, "gfc_realting.xml")
    c["realting"] = len(ET.parse(p).getroot().findall("property")) if os.path.exists(p) else None
    p = os.path.join(d, "gfc_prian.xml")
    c["prian"] = len({o.findtext("id") for o in ET.parse(p).getroot().findall("objects/object")}) if os.path.exists(p) else None
    p = os.path.join(d, "gfc_site.json")
    c["site"] = json.load(open(p, encoding="utf-8"))["count"] if os.path.exists(p) else None
    return c

PRICE = re.compile(r"(?:€|EUR)\s?(\d{1,3}(?:[,.\s\u00a0\u202f]\d{3})+|\d{5,})|(\d{1,3}(?:[,.\s\u00a0\u202f]\d{3})+|\d{5,})\s?(?:€|EUR|евро|євро)")
def prices(t):  # all euro amounts written in a text, as ints
    return {int(re.sub(r"\D", "", a or b)) for a, b in PRICE.findall(t or "")}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--new", required=True); ap.add_argument("--old", required=True)
    ap.add_argument("--root", default="."); ap.add_argument("--log"); ap.add_argument("--allow-drop", default="ні")
    a = ap.parse_args()
    err, warn = [], []
    new, old = counts(a.new), counts(a.old)
    allow = a.allow_drop.strip().lower() in ("так", "yes", "true", "1")
    for k in ("realting", "prian", "site"):
        n, o = new[k], old[k]
        if not n: err.append(f"{k}: 0 оголошень – публікувати не можна (платформа видалила б усе)"); continue
        if o and (o - n > 2 or n < 0.75 * o):
            (warn if allow else err).append(f"{k}: було {o}, стало {n} оголошень" + (" (дозволено вручну)" if allow else
                " – різке зменшення. Якщо так і треба, запустіть вручну з «Дозволити зменшення» = так"))
    base = "https://asistentkaviktorie-crypto.github.io/gfc-feeds/"
    def imgcheck(ref, urls):
        if not urls: err.append(f"{ref}: немає жодного фото")
        for u in urls:
            if u.startswith(base) and not os.path.exists(os.path.join(a.root, u[len(base):])):
                err.append(f"{ref}: файл фото {u[len(base):]} відсутній у репозиторії")
    k = ET.parse(os.path.join(a.new, "gfc_realting.xml")).getroot()
    for p in k.findall("property"):
        ref = p.findtext("ref")
        if not (p.findtext("price") or "").strip().isdigit() or int(p.findtext("price")) <= 0: err.append(f"{ref}: порожня або нульова ціна")
        imgcheck(ref, [u.text for u in p.findall("images/image/url")])
        price = int(p.findtext("price") or 0)
        for L in ("en", "ru"):
            ps = prices(p.findtext(f"desc/{L}"))
            if price and ps and price not in ps:
                warn.append(f"{ref}: у таблиці ціна {price:,} €, а в тексті {L.upper()} (вкладка «Тексти») є лише {sorted(ps)} – перевірте, чи текст не зі старою ціною")
    pr = ET.parse(os.path.join(a.new, "gfc_prian.xml")).getroot()
    for o in pr.findall("objects/object"):
        if not (o.findtext("price") or "").strip().isdigit() or int(o.findtext("price")) <= 0: err.append(f"prian {o.findtext('id')}: порожня ціна")
        imgcheck("prian " + o.findtext("id"), [f.text for f in o.findall("images/image/filename")])
    s = json.load(open(os.path.join(a.new, "gfc_site.json"), encoding="utf-8"))
    for p in s["properties"]:
        if not p.get("price"): err.append(f"сайт {p['ref']}: порожня ціна")
        imgcheck("сайт " + p["ref"], p.get("images") or [])
        for L in s["languages"]:
            ps = prices(" ".join((p["i18n"].get(L) or {}).values()))
            if p.get("price") and ps and p["price"] not in ps:
                warn.append(f"сайт {p['ref']}: у таблиці ціна {p['price']:,} €, а в тексті мовою {L} (вкладка «Переклади») – {sorted(ps)[:4]}")
            if any(not (p["i18n"].get(L, {}).get(f) or "").strip() for f in ("title", "summary", "lifestyle", "invest", "risk")):
                err.append(f"сайт {p['ref']}: неповний текст мовою {L}")
    try:
        import xmlschema
        e = list(xmlschema.XMLSchema(os.path.join(HERE, "kyeroV3.0.xsd")).iter_errors(os.path.join(a.new, "gfc_realting.xml")))
        if e: err.append(f"gfc_realting.xml: {len(e)} помилок формату Kyero (XSD), напр. {str(e[0].reason)[:120]}")
    except ImportError:
        warn.append("xmlschema не встановлено – XSD не перевірено")
    if a.log and os.path.exists(a.log):
        warn += [l[5:].strip() for l in open(a.log, encoding="utf-8") if l.startswith("WARN ") and "jamesedition" not in l]
        warn += [l[6:].strip() for l in open(a.log, encoding="utf-8") if l.startswith("PHOTO ")]
    lines = ["## Оновлення фідів GFC", "", f"Оголошень: Realting {new['realting']} (було {old['realting']}), Prian {new['prian']} (було {old['prian']}), сайт {new['site']} (було {old['site']})", ""]
    lines += (["### ❌ Не опубліковано – виправте в таблиці:"] + [f"- {x}" for x in dict.fromkeys(err)] + [""]) if err else ["### ✅ Перевірка пройдена", ""]
    if warn: lines += ["### ⚠️ Попередження (не блокують):"] + [f"- {x}" for x in dict.fromkeys(warn)]
    txt = "\n".join(lines) + "\n"; print(txt)
    if os.environ.get("GITHUB_STEP_SUMMARY"): open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8").write(txt)
    sys.exit(1 if err else 0)

if __name__ == "__main__":
    main()
