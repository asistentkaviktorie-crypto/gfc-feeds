#!/usr/bin/env python3
"""Green Future Capital – feed builder (self-service version: everything comes from the Google Sheet).

One command regenerates all three feeds from the master sheet (xlsx export of the
Google Sheet "GFC", tab 1 = Об'єкти, tab "Тексти"):

    python3 build_feed.py --xlsx GFC.xlsx --image-base-url https://<host>/images

Outputs (in --out, default ./out):
    gfc_realting.xml      Kyero XML v3  – rows with "У фід Realting" = так
    gfc_jamesedition.xml  Kyero XML v3  – rows with "У фід JamesEdition" = так
    gfc_prian.xml         Prian XML     – rows with "У фід Prian" = так
                                          (empty / formula => same as Realting)

Images: local JPGs in ./images/GFC-<id>/NN.jpg (made by fetch_images.py from the
Realting gallery). With --image-base-url they are referenced as
<base>/GFC-<id>/NN.jpg (upload the images/ folder there). Without it the feeds fall
back to the original Realting .webp URLs (PREVIEW ONLY: fails Kyero XSD and Prian
requires jpg/gif/png).

Dates: feed_state.json keeps a content hash per listing; <date> only changes when the
listing content changes (Kyero/JE update a property only when the date changes).
No Spanish text is generated. Refs are stable: GFC-<RealtingID> (Prian id: gfc<RealtingID>,
because Prian ids must be [a-z0-9]).
"""
import argparse, datetime as dt, hashlib, json, os, re, sys
from xml.sax.saxutils import escape
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)          # repo root: images/ live here
sys.path.insert(0, HERE)
import sheet_config

# Per-listing metadata now lives in the sheet tab «Параметри» (see sheet_config.py); filled in main().
META = {}
PHOTOS = {}   # rid -> [dict(n=<file stem>, w, h, src)] from tab «Фото» + files in images/
MIN_SIDE = 600          # Prian hard minimum; also drops Realting banner strips
KYERO_MIN = (1280, 960) # Kyero recommended minimum (reported only)

YES = {"так", "yes", "y", "1", "true", "да"}

def num(v):
    if v in (None, ""): return None
    try: return float(str(v).replace(",", ".").replace(" ", "").replace("\u00a0", ""))
    except ValueError: return None

def intval(v):
    n = num(v); return None if n is None else int(round(n))

def refid(v):
    """Sheet "Realting ID" -> listing key. Numeric Realting IDs stay as-is ("3987679");
    feed-only listings without a Realting ID use a code, e.g. "GFC-CAPRI" or "CAPRI" -> "CAPRI"."""
    if v in (None, ""): return None
    n = intval(v)
    if n: return str(n)
    s = str(v).strip().upper()
    if s.startswith("GFC-"): s = s[4:]
    return s if re.fullmatch(r"[A-Z0-9]+", s) else None

def load(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=False)
    ws = wb.worksheets[0]
    hdr = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(hdr)}
    need = ["Realting ID", "Ціна €", "У фід Realting", "У фід JamesEdition", "У фід Prian"]
    miss = [h for h in need if h not in col]
    if miss: sys.exit(f"Missing columns in first tab: {miss}")
    tx = {}
    wt = wb["Тексти"]
    th = [str(c.value).strip() if c.value else "" for c in wt[1]]
    for r in wt.iter_rows(min_row=2, values_only=True):
        if r[0] is None: continue
        d = dict(zip(th, r)); k = refid(r[0])
        if k: tx[k] = d
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        g = lambda h: r[col[h]] if h in col and col[h] < len(r) else None
        rid = refid(g("Realting ID"))
        if not rid: continue
        flag = lambda h: str(g(h) or "").strip().lower()
        rl = flag("У фід Realting") in YES
        pr_raw = flag("У фід Prian")
        pr = rl if (pr_raw == "" or pr_raw.startswith("=")) else pr_raw in YES
        rows.append(dict(rid=rid, project=g("Проєкт"), price=intval(g("Ціна €")),
            total=intval(g("Площа заг., м²")), living=intval(g("Житлова, м²")), plot=intval(g("Ділянка, м²")),
            beds=intval(g("Спальні")), baths=num(g("Ванні")), status=g("Realting статус"),
            feeds=dict(realting=rl, jamesedition=flag("У фід JamesEdition") in YES, prian=pr),
            text=tx.get(rid)))
    return rows

def images(rid, base, warn):
    out = []
    for i, m in enumerate(PHOTOS.get(rid, []), 1):
        if min(m["w"], m["h"]) < MIN_SIDE:
            warn.append(f"GFC-{rid}: image {m['n']}.jpg {m['w']}x{m['h']} skipped (<{MIN_SIDE}px)"); continue
        url = f"{base.rstrip('/')}/GFC-{rid}/{m['n']}.jpg" if base else m.get("src")
        out.append((url, m))
    return out[:50]

def stamp(state, ref, payload, now):
    h = hashlib.sha1(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
    s = state.get(ref)
    if not s or s["hash"] != h:
        state[ref] = s = dict(hash=h, date=now)
    return s["date"]

def kx(s):  # Kyero text: escaped, no HTML, line breaks as &#13;
    return escape(s.strip()).replace("\r\n", "\n").replace("\n", "&#13;")

def kyero(rows, base, state, now, warn):
    o = ['<?xml version="1.0" encoding="utf-8"?>', "<root>", "<kyero>", "<feed_version>3</feed_version>",
         f"<feed_generated>{now}</feed_generated>", "</kyero>"]
    for r in rows:
        m, t, rid = META[r["rid"]], r["text"], r["rid"]
        imgs = images(rid, base, warn)
        land = m["kt"] == "plot"
        built = None if land else r["total"]
        beds = 0 if land else (r["beds"] or 0)
        baths = 0 if land else int(r["baths"] or 0)
        desc_en = f'{t["Title EN"]}\n\n{t["Description EN"]}'
        desc_ru = f'{t["Title RU"]}\n\n{t["Description RU"]}'
        payload = [r["price"], built, r["plot"], beds, baths, desc_en, desc_ru, [u for u, _ in imgs], m]
        date = stamp(state, f"GFC-{rid}", payload, now)
        p = ["<property>", f"<id>GFC-{rid}</id>", f"<date>{date}</date>", f"<ref>GFC-{rid}</ref>",
             f"<price>{r['price']}</price>", "<currency>EUR</currency>", "<price_freq>sale</price_freq>",
             f"<new_build>{m['nb']}</new_build>", f"<type>{m['kt']}</type>", f"<town>{m['town']}</town>",
             f"<province>{m['province']}</province>", "<country>Spain</country>",
             "<location>", f"<latitude>{m['lat']}</latitude>", f"<longitude>{m['lon']}</longitude>", "</location>",
             f"<location_detail>{escape(m['detail'])}</location_detail>",
             f"<beds>{beds}</beds>", f"<baths>{baths}</baths>", f"<pool>{m['pool']}</pool>"]
        if built or r["plot"]:
            p.append("<surface_area>")
            if built: p.append(f"<built>{built}</built>")
            if r["plot"]: p.append(f"<plot>{r['plot']}</plot>")
            p.append("</surface_area>")
        if m.get("energy"):
            p += ["<energy_rating>", f"<consumption>{m['energy']}</consumption>", "</energy_rating>"]
        p += ["<desc>", f"<en>{kx(desc_en)}</en>", f"<ru>{kx(desc_ru)}</ru>", "</desc>"]
        if m.get("features"):
            p += ["<features>"] + [f"<feature>{escape(f[:35])}</feature>" for f in m["features"][:50]] + ["</features>"]
        p += ["<images>"]
        for i, (u, _) in enumerate(imgs, 1):
            p += [f'<image id="{i}">', f"<url>{escape(u)}</url>", "</image>"]
        p += ["</images>", "</property>"]
        o += p
    o.append("</root>")
    return "\n".join(o) + "\n"

def para(s):
    return "".join(f"<p>{escape(x.strip())}</p>" for x in re.split(r"\n\s*\n", s.strip()) if x.strip())

def prian(rows, base, state, now, warn):
    o = ["<?xml version='1.0' encoding='utf-8' ?>", "<root>", "<objects>"]
    for r in rows:
        m, t, rid = META[r["rid"]], r["text"], r["rid"]
        imgs = images(rid, base, warn)
        if not imgs: warn.append(f"GFC-{rid}: no usable images – Prian will reject it")
        land = m["pt"] == 10
        for lang, (ttl, dsc, st, addr) in ((1, ("Title RU", "Description RU", m["st_ru"], m["addr_ru"])),
                                           (3, ("Title EN", "Description EN", m["st_en"], m["addr_en"]))):
            if len(st) > 40: warn.append(f"GFC-{rid}: specialtxt >40 chars: {st}")
            p = ["<object>", f"<id>gfc{rid.lower()}</id>", f"<language_id>{lang}</language_id>", f"<object_id>{m['pt']}</object_id>",
                 "<type_id>1</type_id>", "<country_id>43</country_id>", f"<district_id>{m['district_id']}</district_id>",
                 f"<address>{escape(addr)}</address>", f"<longitude>{m['lon']}</longitude>", f"<latitude>{m['lat']}</latitude>",
                 "<currency_id>2</currency_id>", f"<price>{r['price']}</price>"]
            if not land:
                if r["beds"]: p.append(f"<bedrooms>{r['beds']}</bedrooms>")
                if r["baths"]: p.append(f"<bathrooms>{int(r['baths'])}</bathrooms>")
                if r["total"]: p.append(f"<square>{r['total']}</square>")
            if r["plot"]: p += [f"<square_land>{r['plot']}</square_land>", "<square_land_unit>129</square_land_unit>"]
            if m["bt"]: p.append(f"<building_type>{m['bt']}</building_type>")
            if m["bt"] and m.get("bdate"): p.append(f"<building_date>{escape(m['bdate'])}</building_date>")
            p += [f"<specialtxt>{escape(st[:40])}</specialtxt>",
                  f"<description><![CDATA[<p><b>{escape(t[ttl])}</b></p>{para(t[dsc])}]]></description>", "<images>"]
            for i, (u, _) in enumerate(imgs):
                p += ["<image>", f"<filename>{escape(u)}</filename>", f"<iorder>{i}</iorder>", "</image>"]
            p += ["</images>", "</object>"]
            o += p
    o += ["</objects>", "</root>"]
    return "\n".join(o) + "\n"

def validate(path, kind, base, warn):
    import xml.etree.ElementTree as ET
    root = ET.parse(path).getroot()  # well-formedness
    if kind == "kyero":
        props = root.findall("property")
        for p in props:
            for f in ["id", "date", "ref", "price", "price_freq", "type", "town", "province", "desc/en", "images"]:
                if p.find(f) is None or (f != "images" and not (p.find(f).text or "").strip()):
                    warn.append(f"{os.path.basename(path)} {p.findtext('ref')}: missing {f}")
            if len(p.findall("images/image")) == 0: warn.append(f"{p.findtext('ref')}: no images")
        try:
            import xmlschema
            xs = xmlschema.XMLSchema(os.path.join(HERE, "kyeroV3.0.xsd"))
            errs = list(xs.iter_errors(path))
            if errs:
                kinds = sorted({e.reason.split(".")[0][:90] if e.reason else str(e)[:90] for e in errs})
                warn.append(f"{os.path.basename(path)}: XSD {len(errs)} error(s): {kinds[:3]}")
            else:
                print(f"  XSD OK: {os.path.basename(path)}")
        except ImportError:
            warn.append("xmlschema not installed – XSD check skipped (pip install xmlschema)")
        if not props: warn.append(f"{os.path.basename(path)}: 0 listings (do not connect an empty feed – the platform would remove everything)")
        return len(props)
    objs = root.findall("objects/object")
    for o in objs:
        for f in ["id", "language_id", "object_id", "type_id", "country_id", "district_id", "currency_id", "price"]:
            if not (o.findtext(f) or "").strip(): warn.append(f"prian {o.findtext('id')}: missing {f}")
        if not re.fullmatch(r"[a-z0-9]+", o.findtext("id") or ""): warn.append(f"prian id not [a-z0-9]: {o.findtext('id')}")
        for fn in o.findall("images/image/filename"):
            if not re.search(r"\.(jpe?g|gif|png)$", fn.text, re.I):
                warn.append(f"prian {o.findtext('id')}: image not jpg/gif/png ({fn.text[-30:]})"); break
    if not objs: warn.append(f"{os.path.basename(path)}: 0 listings (do not connect an empty feed – JE/Kyero would remove everything)")
    return len({o.findtext("id") for o in objs})

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default="/workspace/gfc_sheet/GFC.xlsx")
    ap.add_argument("--out", default=os.path.join(HERE, "out"))
    ap.add_argument("--image-base-url", default=os.environ.get("GFC_IMAGE_BASE_URL"))
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    stf = os.path.join(HERE, "feed_state.json")
    state = json.load(open(stf)) if os.path.exists(stf) else {}
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    warn = []
    rows = load(a.xlsx)
    cfg = sheet_config.load(a.xlsx, os.path.join(ROOT, "images"), warn)
    META.update(cfg["META"]); PHOTOS.update(cfg["PHOTOS"])
    if not a.image_base_url:
        warn.append("No --image-base-url: using Realting .webp URLs (PREVIEW ONLY, not valid for Kyero XSD/Prian)")
    sel = {}
    for feed in ("realting", "jamesedition", "prian"):
        s = []
        for r in rows:
            if not r["feeds"][feed]: continue
            if r["rid"] not in META: warn.append(f"{feed}: {r['rid']} flagged but has no row in tab «Параметри» – skipped"); continue
            if r["rid"].isdigit() and r["status"] != "Active": warn.append(f"{feed}: {r['rid']} flagged but Realting status = {r['status']}")
            if not r["text"]: warn.append(f"{feed}: {r['rid']} has no texts – skipped"); continue
            if feed == "jamesedition" and (r["price"] or 0) < 450000:
                warn.append(f"jamesedition: GFC-{r['rid']} €{r['price']:,} is below JE minimum €450,000")
            s.append(r)
        sel[feed] = s
    # Kyero state is shared (same content => same date in both Kyero feeds)
    files = {"realting": ("gfc_realting.xml", "kyero"), "jamesedition": ("gfc_jamesedition.xml", "kyero"), "prian": ("gfc_prian.xml", "prian")}
    iw = []
    for feed, (fn, kind) in files.items():
        xml = kyero(sel[feed], a.image_base_url, state, now, iw) if kind == "kyero" else prian(sel[feed], a.image_base_url, state, now, iw)
        path = os.path.join(a.out, fn)
        open(path, "w", encoding="utf-8").write(xml)
        n = validate(path, kind, a.image_base_url, warn)
        print(f"{fn}: {n} listings -> {path}")
    json.dump(state, open(stf, "w"), indent=1, ensure_ascii=False)
    # website feed (gfc_site.json) – separate file, does not affect the XML feeds
    import build_site
    build_site.build(rows, images, a.image_base_url, a.out, warn, cfg)
    for w in sorted(set(iw)) + list(dict.fromkeys(warn)): print("WARN", w)

if __name__ == "__main__":
    main()
