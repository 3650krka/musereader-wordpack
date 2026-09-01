import json, sys, collections, pathlib

sys.stdout.reconfigure(encoding="utf-8")
data = pathlib.Path(__file__).resolve().parent.parent / "musereader-wordpack-graded.json"
entries = json.loads(data.read_text(encoding="utf-8"))
if not isinstance(entries, list):
    entries = entries.get("entries") or entries.get("words") or []
print("total_entries", len(entries))
levels = collections.Counter(e.get("level") for e in entries)
order = ["zhongKao", "gaoKao", "cet4", "cet6", "kaoYan", "ielts", "toefl", "tem4", "tem8", "gre"]
for key in order:
    print(f"  {key}: {levels.get(key, 0)}")
extra = {k: v for k, v in levels.items() if k not in order}
if extra:
    print("  UNEXPECTED_LEVELS:", extra)


def filled(entry, field):
    value = entry.get(field)
    return bool(value) and str(value).strip() not in ("", "无")


for field in ("phonetic", "definition", "translation", "exchange", "collins", "oxford", "bnc", "root", "tag"):
    count = sum(1 for e in entries if filled(e, field))
    print(f"with_{field}: {count} ({count * 100 // max(1, len(entries))}%)")
