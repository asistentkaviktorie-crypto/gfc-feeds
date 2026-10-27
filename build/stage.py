#!/usr/bin/env python3
"""Copy the new feeds into the repo root only if their content changed (timestamps feed_generated /
generated are ignored, so a day without edits makes no commit). Prints CHANGED or UNCHANGED."""
import json, os, re, shutil, sys
new, root = sys.argv[1], sys.argv[2]
def norm(p):
    if not os.path.exists(p): return None
    t = open(p, encoding="utf-8").read()
    if p.endswith(".json"):
        d = json.loads(t); d.pop("generated", None); return json.dumps(d, sort_keys=True)
    return re.sub(r"<feed_generated>.*?</feed_generated>", "", t)
changed = False
for f in ("gfc_realting.xml", "gfc_prian.xml", "gfc_site.json"):
    if norm(os.path.join(new, f)) != norm(os.path.join(root, f)):
        shutil.copy2(os.path.join(new, f), os.path.join(root, f)); changed = True; print("changed:", f)
print("CHANGED" if changed else "UNCHANGED")
