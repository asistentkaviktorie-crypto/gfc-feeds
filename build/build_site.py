"""Build gfc_site.json – listings feed for the GFC website (bloom-coral-crane-rapid.grok.me).
Numbers (price/sqm/beds/baths) come from tab «Обєкти», photos from tab «Фото» (same URLs as the XML feeds),
site-only fields from tab «Сайт», texts from tab «Переклади» (all via sheet_config.py).
Called from build_feed.py main(); does not touch the XML feeds."""
import json, os, sys, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
LANGS = ["en", "es", "ru", "de", "fr", "cs"]
FIELDS = ["title", "summary", "lifestyle", "invest", "risk"]

def _n(v):
    if v is None: return None
    v = float(v); return int(v) if v.is_integer() else v

def build(rows, images, base, out_dir, warn, cfg):
    TXT, COVER = cfg["TXT"], cfg["COVER"]
    by = {r["rid"]: r for r in rows}
    props = []
    for m in cfg["SITE"]:
        k = m["key"]; r = by.get(k)
        if r is None:
            raise SystemExit(f"ПОМИЛКА сайт: GFC-{k} є у вкладці «Сайт», але немає в «Обєкти»")
        land = m["type"] == "land"
        price = r.get("price")
        if not price: raise SystemExit(f"ПОМИЛКА сайт: GFC-{k} – порожня «Ціна €» у вкладці «Обєкти»")
        sqm = (r.get("plot") if land else r.get("total"))
        beds = None if land else r.get("beds")
        baths = None if land else r.get("baths")
        imgs = [u for u, _ in images(k, base, [])]
        c = COVER.get(k)
        if c:
            first = [u for u in imgs if u.endswith("/" + c)]
            imgs = first + [u for u in imgs if u not in first]
        if not imgs: warn.append(f"site: {k} has no images")
        i18n = {}
        for L in LANGS:
            t = TXT[L].get(k)
            if not t or any(not (t.get(f) or "").strip() for f in FIELDS):
                raise SystemExit(f"ПОМИЛКА сайт: GFC-{k} – немає повного тексту мовою {L} у вкладці «Переклади»")
            i18n[L] = {f: t[f].strip() for f in FIELDS}
        props.append(dict(
            id=m["id"], cabinetId=m["cabinetId"], ref=f"GFC-{k}", type=m["type"], area=m["area"], town=m["town"],
            price=_n(price), priceFrom=m["priceFrom"], currency="EUR", sqm=_n(sqm), beds=_n(beds), baths=_n(baths),
            seaM=m["seaM"], license=m["license"], communityMonth=m["communityMonth"], ibiYear=m["ibiYear"],
            insuranceYear=m["insuranceYear"], longRentMonth=m["longRentMonth"], adr=m["adr"],
            occLong=m["occLong"], occShort=m["occShort"], mgmtLong=m["mgmtLong"], mgmtShort=m["mgmtShort"],
            landUse=None, sellableSqm=None, buildCostSqm=None, exitSqm=None,
            contract=False, hidden=m["hidden"], images=imgs, i18n=i18n))
    if len({p["id"] for p in props}) != len(props):
        raise SystemExit("ПОМИЛКА сайт: однаковий «ID на сайті» у двох рядках вкладки «Сайт»")
    doc = dict(version=1, generated=dt.datetime.now().astimezone().isoformat(timespec="seconds"),
               currency="EUR", languages=LANGS, fallbackLanguages=["en", "ru"],
               count=len(props), visibleCount=sum(not p["hidden"] for p in props), properties=props)
    path = os.path.join(out_dir, "gfc_site.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=1); f.write("\n")
    json.load(open(path, encoding="utf-8"))
    print(f"gfc_site.json: {doc['count']} properties ({doc['visibleCount']} visible, {doc['count']-doc['visibleCount']} hidden) -> {path}")
    return doc["count"]
