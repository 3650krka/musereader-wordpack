# MuseReader Open Word Pack（开放分级词库包）

`musereader-wordpack-graded.json`：**42,312** 个英文单词的十档分级词库，供 [MuseReader](../click/musetranslate/) 的 WordWise 生词透析 / 背诵词级**导入使用**（词库不随应用分发）。

## 档位分布

| 档 | 词数 | | 档 | 词数 |
|---|---|---|---|---|
| 中考 | 8,272 | | 托福 | 2,428 |
| 高考 | 4,763 | | 专四 | 808 |
| CET4 | 7,741 | | 专八 | 9,661 |
| CET6 | 1,011 | | GRE | 2,157 |
| 考研 | 2,636 | | 合计 | **42,312** |
| 雅思 | 2,835 | | | |

一词多档取**最低档**（与 MuseReader Rust 端 `WordLevelIndex` 同语义）。

## 字段

`word` / `level`（必填）+ 可选：`phonetic`（37,091 / 87.7%）、`definition`（41,796 / 98.8%）、`translation`（40,816 / 96.5%，ECDICT 完整中文义）、`tag`（考试标签）、`exchange`（23,146 / 54.7%，词形变化 `d:/p:/i:/3:/s:` 编码）、`collins`（12,556，1–5 星级）、`oxford`（牛津3000）、`bnc` / `frq`（词频排名，27,116）、`root`（词根助记，5,475 条）。

空字段直接省略以控制体积（单文件 ≈ 10 MB）。

## 生成方式（可复现）

```bash
pip install -r requirements.txt        # 仅需 openpyxl
# 素材克隆到与本仓库同级的 reference-products/（或用参数显式指定路径）：
#   git clone https://github.com/lilinji/English   reference-products/English
#   git clone https://github.com/skywind3000/ECDICT reference-products/ecdict
python tools/build_wordpack_open.py [English仓库路径] [ECDICT目录] [输出json]

# 校验词条数与档位分布：
python tools/stats.py
```

`tools/build_wordpack_open.py` 复用同目录 `tools/build_word_levels.py`（共用 `WORD_RE`、归一化与合并语义，与 MuseReader Rust 端 `WordLevelIndex` 同规则），两个文件需一起保留。

素材构成（全量，非抽样）：
1. **lilinji/English** 词表仓库 —— 全部 949 份 xlsx（10 档大纲/考研/雅思/托福/专四/专八/GRE + 全国 20 个教材版本中小学同步词书；短语/句型表按设计排除）；
2. **skywind3000/ECDICT** `ecdict.csv` —— 中文翻译、考试标签、柯林斯/牛津、BNC/FRQ 词频、词形变换（含 198 个标签反哺词）；
3. **ECDICT wordroot.txt** —— 词根助记表。

难度归属采用**结构化推导**：词表所在目录与文件名关键词 → 档位；不使用任何逐词白名单/黑名单。

## 导入到 MuseReader

设置 → 词库 → 导入词库 → 选择本目录的 `musereader-wordpack-graded.json`。
解析 <5 秒；导入后行间生词标注、词卡、背诵词级全部激活。

## 许可与署名（重要）

完整法律条款见 [`LICENSE`](LICENSE)（CC BY-NC-SA 4.0 官方全文），上游素材逐项署名见 [`NOTICE.md`](NOTICE.md)。

- 源词表仓库 [lilinji/English](https://github.com/lilinji/English)：**CC BY-NC-SA 4.0**（署名 + **非商业性使用** + 相同方式共享），且其 README 声明"词书版权归原作者/出版社所有"。
- ECDICT：MIT（词典数据为公开汇编）。
- 因此本衍生包整体按 **CC BY-NC-SA 4.0** 提供：**可自由使用、修改、分享，但不得商用，须署名上游来源**。MuseReader 应用本体不内嵌本包，导入行为由用户完成；若将本包纳入任何商业分发，请先解决上游授权。
