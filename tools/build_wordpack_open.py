#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""全量分级词库生成器（开放词包版）——把 reference-products 的全部开源素材合成单一可分发词包。

与 build_word_levels.py 的区别：
  * 旧版每档只取 1 份样表（共 10 份）；本版扫 English 仓库 **全部 949 份 xlsx**。
  * 难度档由 **目录层级 + 文件名关键词** 结构化推导（不逐词判断，不依赖词面特例）。
  * 叠加 ECDICT **考试标签反哺**：xlsx 未覆盖但 ecdict 带 zk/gk/cet4…/gre 标签的词直接入包。
  * 全部词条统一用 ECDICT 富化（translation/tag/exchange/collins/oxford/bnc/frq）+ 词根表。

产物（单独开源包，不进应用仓库）：
  D:/Projects/musetranslate/wordpack-open/musereader-wordpack-graded.json  +  生成统计打印

用法：python tools/build_wordpack_open.py [English仓库路径] [ECDICT目录] [输出json]
"""

from __future__ import annotations

import csv
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_word_levels as base  # 复用：WORD_RE/归一化/合并语义（与 Rust 端同规则）

ROOT = Path(__file__).resolve().parent.parent
ENGLISH = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT.parent.parent / "reference-products" / "English"
ECDICT_DIR = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT.parent.parent / "reference-products" / "ecdict"
OUT = Path(sys.argv[3]) if len(sys.argv) > 3 else ROOT.parent.parent / "wordpack-open" / "musereader-wordpack-graded.json"

# 顶层目录 → 难度档（结构化：目录即大纲归属）
DIR_LEVELS = {
    "1.全国各大教材版本中小学同步": "zhongKao",
    "2.中考": "zhongKao",
    "2.高考": "gaoKao",
    "3.专四": "tem4",
    "3.四级": "cet4",
    "3.大学英语": "cet4",
    "4.专八": "tem8",
    "4.六级": "cet6",
    "5.考研": "kaoYan",
    "6.研究生": "kaoYan",
    "6.考博": "tem8",
    "7.托福": "toefl",
    "7.雅思": "ielts",
    "8.商务英语": "cet6",
    "8.国际英语": "cet6",
    "8.新世纪英专": "tem4",
    "8.新概念英语": "cet4",  # 册级细分见 FILE_KEYWORDS
    "9.其他（更多）": "tem8",  # 未知主题表按高档兜底（高档=少标注，绝不把生词标成低档）
    "9.扇贝英语（IT）": "tem8",
}

# 文件名关键词（大小写不敏感）→ 难度档；先匹配先赢，排在具体目录兜底之前。
FILE_KEYWORDS: list[tuple[re.Pattern[str], str]] = [
    (re.compile(p, re.IGNORECASE), lv)
    for p, lv in [
        (r"GRE|GMAT|巴朗|Barron|SAT|ACT", "gre"),
        (r"^[0-9１-９]?\s*[～~—-]\s*[0-9１-９六四]?\s*级|^1[～~—-]6|1～6级", "zhongKao"),  # 分级总表含入门级 → 最低档
        (r"考博|PHD", "tem8"),
        (r"四六级|四、?六", "cet4"),  # 四六混合表按最低档（取低不取高）
        (r"专业[四4４]级|专四|专[4４]|TEM-?4", "tem4"),
        (r"专业[八8８]级|专八|专[8８]|八级|TEM-?8", "tem8"),
        (r"新概念.*第四册", "tem4"),
        (r"新概念.*第三册", "cet6"),
        (r"新概念.*第二册", "cet4"),
        (r"新概念.*第一册", "gaoKao"),
        (r"雅思|IELTS", "ielts"),
        (r"托福|TOEFL", "toefl"),
        (r"考研|POSTGRAD", "kaoYan"),
        (r"六级|[6６六]\s*级", "cet6"),
        (r"四级|[4４四]\s*级|CET-?4|FCE", "cet4"),
        (r"PET|商务英语|BEC", "cet4"),
        (r"高考|高中|高[一二三123]$|高[一二三123]册?|高一上|高一下|高二上|高二下|高三上|高三下", "gaoKao"),
        (r"中考|初中|小学|KET|YLE|Movers|Flyers|Starters", "zhongKao"),
    ]
]

# ECDICT tag 词元 → 难度档（与 Rust level_from_tag 对齐 + 补充）
TAG_LEVELS = {
    "zk": "zhongKao", "gk": "gaoKao", "cet4": "cet4", "cet6": "cet6", "ky": "kaoYan",
    "ielts": "ielts", "toefl": "toefl", "tem4": "tem4", "tem8": "tem8",
    "gre": "gre", "sat": "gre", "gmat": "gre",
}


def level_for(path: Path) -> str:
    """目录定位档 → 文件名关键词覆盖档（结构信号优先于文件名兜底）。"""
    rel_parts = path.relative_to(ENGLISH).parts
    top = rel_parts[0] if rel_parts else ""
    level = DIR_LEVELS.get(top, "tem8")
    name = path.name
    for pattern, mapped in FILE_KEYWORDS:
        if pattern.search(name):
            return mapped
    if top == "1.全国各大教材版本中小学同步":
        return "zhongKao"  # 教材同步表名不含关键词时保持目录档
    return level


def main() -> None:
    if not ENGLISH.is_dir():
        sys.exit(f"English 仓库不存在: {ENGLISH}")
    xlsx_files = sorted(ENGLISH.rglob("*.xlsx"))
    xlsx_files = [p for p in xlsx_files if not p.name.startswith("~$")]
    print(f"发现 xlsx {len(xlsx_files)} 份", flush=True)

    entries: list[dict] = []
    per_level: Counter[str] = Counter()
    failed: list[str] = []
    for i, path in enumerate(xlsx_files):
        level = level_for(path)
        try:
            rows = base.load_sheet(path, level)
        except Exception as error:  # noqa: BLE001 单表损坏不阻断全量
            failed.append(f"{path.name}: {str(error)[:80]}")
            continue
        if not rows:
            failed.append(f"{path.name}: 0 行有效")
            continue
        entries.extend(rows)
        per_level[level] += len(rows)
        if (i + 1) % 100 == 0:
            print(f"  …{i + 1}/{len(xlsx_files)} 文件，累计 {len(entries)} 行", flush=True)
    print(f"xlsx 解析完成：{len(entries)} 行（失败 {len(failed)}）", flush=True)
    for msg in failed[:12]:
        print("  skip:", msg)

    merged = base.merge_levels(entries)
    print(f"去重后独立词: {len(merged)}", flush=True)

    # ECDICT 标签反哺：把 xlsx 未覆盖、但大纲标签明确的词并入（取标签最低档）。
    if not (ECDICT_DIR / "ecdict.csv").is_file():
        sys.exit(f"ECDICT 缺失: {ECDICT_DIR / 'ecdict.csv'}")
    bonus = 0
    unknown_tags: Counter[str] = Counter()
    with open(ECDICT_DIR / "ecdict.csv", newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            w = (row.get("word") or "").strip().lower()
            if not w or len(w) > base.MAX_WORD_LEN or not base.WORD_RE.match(w) or " " in w:
                continue
            tag = (row.get("tag") or "").strip()
            if not tag:
                continue
            levels = []
            for token in tag.split():
                mapped = TAG_LEVELS.get(token)
                if mapped:
                    levels.append(mapped)
                else:
                    unknown_tags[token] += 1
            if not levels:
                continue
            best = min(levels, key=lambda lv: base.LEVEL_ORDER[lv])
            existing = merged.get(w)
            if existing is None:
                merged[w] = {"word": w, "level": best, "phonetic": "", "definition": ""}
                bonus += 1
            elif base.LEVEL_ORDER[best] < base.LEVEL_ORDER[existing["level"]]:
                existing["level"] = best
    print(f"ECDICT 标签反哺新增 {bonus} 词；未知标签词元 {dict(unknown_tags.most_common(10))}", flush=True)

    ecdict_hits = base.merge_ecdict(merged)
    root_hits = base.merge_wordroot(merged)

    out_entries = [merged[w] for w in sorted(merged)]
    final_levels: Counter[str] = Counter(e["level"] for e in out_entries)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(out_entries, ensure_ascii=False, separators=(",", ":"))
    OUT.write_text(payload, encoding="utf-8")
    print(f"OK -> {OUT}")
    print(f"entries={len(out_entries):,} bytes={len(payload):,} ecdict富化={ecdict_hits:,} 词根={root_hits:,}")
    print("档位分布:", json.dumps(dict(sorted(final_levels.items(), key=lambda kv: base.LEVEL_ORDER[kv[0]])), ensure_ascii=False))


if __name__ == "__main__":
    main()
