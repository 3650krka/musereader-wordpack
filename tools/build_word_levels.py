#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 lilinji/English 词表 xlsx 与 ECDICT 附加数据合成为 MuseTranslate 内置 word_levels.json。

词源：
  1. reference-products/English（lilinji/English 克隆）—— 难度档、音标、基础释义。
  2. reference-products/ecdict/ecdict.csv（skywind3000/ECDICT）—— 完整中文翻译、
     考试标签(tag)、柯林斯星级(collins)、牛津3000(oxford)、BNC/词频排名(bnc/frq)、
     词形变换(exchange)。
  3. reference-products/ecdict/wordroot.txt（ECDICT 词根表，JSON）—— 词根助记。

每个难度档选一个"大纲/正序"权威词表；一词多档取最低档（与 Rust 端
WordLevelIndex::from_entries 同语义），空字段由其他档/词源补全。

产物：src-tauri/assets/word_levels.json（可选字段为空时省略以控制体积）：
  [{"word":"abandon","level":"cet4","phonetic":"/əˈbændən/","definition":"v. 放弃；抛弃",
    "translation":"vt. 放弃, 抛弃, ...","tag":"gk cet4 cet6 ky toefl gre","collins":3,
    "oxford":1,"bnc":2057,"frq":2182,"exchange":"d:abandoned/p:abandoned/...",
    "root":"ab：离开（Latin）"}, ...]

level 取值与 word_levels.rs 的 serde(camelCase) 对应：
  zhongKao/gaoKao/cet4/cet6/kaoYan/ielts/toefl/tem4/tem8/gre

用法：python tools/build_word_levels.py [English仓库路径] [ECDICT目录]
"""

from __future__ import annotations

import csv
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
ENGLISH = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT.parent.parent / "reference-products" / "English"
ECDICT_DIR = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT.parent.parent / "reference-products" / "ecdict"
ECDICT_CSV = ECDICT_DIR / "ecdict.csv"
WORDROOT = ECDICT_DIR / "wordroot.txt"
OUT = ROOT / "src-tauri" / "assets" / "word_levels.json"

# (难度档, 相对 English 仓库的词表文件) —— 每档选"大纲/正序/精选"覆盖最全的一份
SOURCES = [
    ("zhongKao", "2.中考/初中英语大纲词汇.xlsx"),
    ("gaoKao", "2.高考/高考英语大纲词汇表.xlsx"),
    ("cet4", "3.四级/2013四级词汇正序版.xlsx"),
    ("cet6", "4.六级/2014六级词汇乱序版.xlsx"),
    ("kaoYan", "5.考研/2013考研词汇正序版.xlsx"),
    ("ielts", "7.雅思/2013雅思词汇乱序版.xlsx"),
    ("toefl", "7.托福/570个单词轻松征服托福.xlsx"),
    ("tem4", "3.专四/专四词汇乱序版.xlsx"),
    ("tem8", "4.专八/2010专八词汇正序版.xlsx"),
    ("gre", "9.其他（更多）/GRE词汇精选乱序版.xlsx"),
]

# 与 word_levels.rs 的 WordLevel 枚举顺序一致（低 → 高）
LEVEL_ORDER = {name: i for i, name in enumerate(
    ["zhongKao", "gaoKao", "cet4", "cet6", "kaoYan", "ielts", "toefl", "tem4", "tem8", "gre"])}

WORD_RE = re.compile(r"^[A-Za-z][A-Za-z'’\-]*$")
MAX_WORD_LEN = 32
MAX_DEF_LEN = 600   # 释义/翻译截断，避免个别超长单元格撑爆体积
MAX_TAG_LEN = 80
MAX_EXCHANGE_LEN = 200
MAX_ROOT_LABEL = 60


def norm_phonetic(uk: object, us: object) -> str:
    raw = (str(us).strip() if us else "") or (str(uk).strip() if uk else "")
    if not raw:
        return ""
    raw = raw.strip("[]")
    return f"/{raw}/" if raw else ""


def norm_definition(cell: object) -> str:
    if cell is None:
        return ""
    text = str(cell).replace("\r", "").strip()
    text = re.sub(r"\n{2,}", "\n", text)
    return text[:MAX_DEF_LEN]


def load_sheet(path: Path, level: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.worksheets[0]
    out: list[dict] = []
    skipped = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:  # 表头：单词/英音/美音/释义
            continue
        if not row or row[0] is None:
            continue
        word = str(row[0]).strip().strip('"').strip()
        if not word or len(word) > MAX_WORD_LEN or not WORD_RE.match(word):
            skipped += 1
            continue
        # 词组/短语（含空格）不进单词索引——WordWise 标注与透析都以单词为单位
        if " " in word:
            skipped += 1
            continue
        uk = row[1] if len(row) > 1 else None
        us = row[2] if len(row) > 2 else None
        definition = row[3] if len(row) > 3 else None
        out.append(
            {
                "word": word.lower(),
                "level": level,
                "phonetic": norm_phonetic(uk, us),
                "definition": norm_definition(definition),
            }
        )
    wb.close()
    print(f"  {level:9s} {path.name}: {len(out)} words (skipped {skipped})", flush=True)
    return out


def merge_levels(entries: list[dict]) -> dict[str, dict]:
    """一词多档取最低档；空音标/释义由其他档补全（不丢已有信息）。"""
    merged: dict[str, dict] = {}
    for e in entries:
        cur = merged.get(e["word"])
        if cur is None:
            merged[e["word"]] = dict(e)
            continue
        if LEVEL_ORDER[e["level"]] < LEVEL_ORDER[cur["level"]]:
            winner, loser = dict(e), cur
            merged[e["word"]] = winner
        else:
            winner, loser = cur, e
        for k in ("phonetic", "definition"):
            if not winner[k] and loser[k]:
                winner[k] = loser[k]
    return merged


def merge_ecdict(merged: dict[str, dict]) -> int:
    """ECDICT 增量合并：翻译/标签/星级/词频/词形（原表已有的难度档与音标不动）。"""
    if not ECDICT_CSV.is_file():
        print(f"!! 缺失 {ECDICT_CSV}，跳过 ECDICT 增强", file=sys.stderr)
        return 0
    hit = 0
    with open(ECDICT_CSV, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            w = (row.get("word") or "").strip().lower()
            if w not in merged:
                continue
            e = merged[w]
            translation = (row.get("translation") or "").replace("\r", "").strip()
            if translation:
                e["translation"] = translation[:MAX_DEF_LEN]
            tag = (row.get("tag") or "").strip()
            if tag:
                e["tag"] = tag[:MAX_TAG_LEN]
            exchange = (row.get("exchange") or "").strip()
            if exchange:
                e["exchange"] = exchange[:MAX_EXCHANGE_LEN]
            for k in ("collins", "oxford", "bnc", "frq"):
                v = (row.get(k) or "").strip()
                if v.isdigit() and int(v) > 0:
                    e[k] = int(v)
            hit += 1
    return hit


def merge_wordroot(merged: dict[str, dict]) -> int:
    """词根反向索引：wordroot.txt 的 example 词命中索引 → 记录词根标签（最多 2 条）。"""
    if not WORDROOT.is_file():
        print(f"!! 缺失 {WORDROOT}，跳过词根合并", file=sys.stderr)
        return 0
    data = json.loads(WORDROOT.read_text(encoding="utf-8"))
    word_roots: dict[str, list[str]] = {}
    for root, info in data.items():
        meaning = str(info.get("meaning") or "").strip()
        if not meaning:
            continue
        origin = str(info.get("origin") or "").strip()
        label = f"{root}：{meaning[:MAX_ROOT_LABEL]}"
        if origin:
            label += f"（{origin[:20]}）"
        for ex in info.get("example") or []:
            w = str(ex).strip().lower()
            if w in merged:
                word_roots.setdefault(w, []).append(label)
    for w, labels in word_roots.items():
        merged[w]["root"] = "；".join(labels[:2])
    return len(word_roots)


def main() -> None:
    if not ENGLISH.is_dir():
        sys.exit(f"English 仓库不存在: {ENGLISH}")
    entries: list[dict] = []
    for level, rel in SOURCES:
        path = ENGLISH / rel
        if not path.is_file():
            print(f"!! 缺失词表 {rel}，跳过该档", file=sys.stderr)
            continue
        entries.extend(load_sheet(path, level))
    merged = merge_levels(entries)
    ecdict_hits = merge_ecdict(merged)
    root_hits = merge_wordroot(merged)
    out_entries = [merged[w] for w in sorted(merged)]
    OUT.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(out_entries, ensure_ascii=False, separators=(",", ":"))
    OUT.write_text(payload, encoding="utf-8")
    print(
        f"OK -> {OUT}  entries={len(out_entries)}  bytes={len(payload):,}  "
        f"ecdict={ecdict_hits}  wordroot={root_hits}"
    )


if __name__ == "__main__":
    main()
