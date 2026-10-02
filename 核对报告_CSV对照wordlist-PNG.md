# merged_vocabulary.csv 对照 word-list PNG 核对报告

核对日期：2026-07-13

## 方法

以各年级 word-list PNG（7a/7b/8a/8b/9a/9b 目录）为标准答案。
逐张图片转写为结构化文本（单词 / 音标 / 词性释义 / 页码）。
再用脚本与 merged_vocabulary.csv 机器比对。
标记含义：🔴 图片有而 CSV 无（遗漏）；🔵 CSV 有而图片无（多余）；🟠 同词但字段不一致。

转写与比对的中间文件都在本次会话的 scratchpad 目录，脚本为 diff_vocab.py，各册转写为 7A.tsv ... 9B.tsv。

## 一句话结论

问题几乎全部集中在 7A 和 8A。
原因是 CSV 各册来自不同来源，只有 8B/9A/9B 直接来自这些 word-list PNG。
7A 最严重，建议整册重做。

| 年级 | CSV 来源 | 与 PNG 一致性 | 关键问题 |
|---|---|---|---|
| 7A | v2024_Vocabulary_7_1.xlsx | ❌ 差 | 音标系统性损坏 + 遗漏19 + 多余4 + 页码错位 + 错别字 |
| 7B | 7b-vocab_1-8.jpg | ✅ 好 | 仅 5 处微小差异 |
| 8A | textbook_..._8_1.md | ⚠️ 一处整块遗漏 | 漏 12 个连续词 + emotional intelligence |
| 8B | 8b-word-list PNG | ✅ 完全一致 | 无 |
| 9A | 9a-word-list PNG | ✅ 近乎一致 | 仅漏 1 词 widely |
| 9B | 9b-word-list PNG | ✅ 一致 | 无实质错误 |

---

## 7A —— 问题严重，建议整册重做

图片 333 条，CSV 318 条。

### 1. 音标系统性损坏（最严重，影响全册）

CSV 中所有含小写字母 l 的音标，l 被误存成大写 I。
例：spell 存成 [speI]，bell 存成 [beI]，Helen 存成 ['heIən]，full 存成 [fʊI]，laugh 存成 [Iɑ:f]，hall 存成 [hɔ:I]，field 存成 [fi:Id]，large 存成 [Iɑ:dʒ]，special 存成 ['speʃI]，classmate 存成 ['kIɑ:smeɪt]，Lily 存成 ['IɪIi]。
这是源头 OCR 把字母 l 认成 I，凡含 l 的词几乎都错，数量达几十个。

### 2. 遗漏 19 个词（图片有，CSV 无）

WHO、circle、page、re、the Great Wall、member、grandparent、centre、different from、send、quarter、at weekends、part、prepare sth for、celebration、contact、symbol、take photos、take a photo。
（其中 take photos 与 take a photo 被 CSV 合并成一条 "take photos / take a photo"。）

### 3. 多余 4 个词（CSV 有，图片无）

CD（图片此处应为 WHO）、pink、White、date。

### 4. 页码错位（图片 → CSV）

each / other / each other：p.2 → p.20。
everyone：p.2 → p.47。
oh：p.2 → p.4。
spell：p.4 → p.2。
chess / Chinese chess：p.30 → p.31。
Jim：p.28 → p.31。
Kate：p.31 → p.32。
（xlsx 把 Starter 阶段的词并到了正式单元，页码体系与新教材 PNG 不同。）

### 5. 释义错误或截断

would：词性写成 model v.（应为 modal v.），释义还被截断。
tennis：图片"网球运动"，CSV 只有"网球"。
across：图片"在对面；横过"，CSV 变成"过；穿过"。
gym / past / across：CSV 释义被截断。

---

## 8A —— 有一处整块遗漏

图片 551 条，CSV 540 条。

### 1. 整块遗漏（重点）：Unit 3 第 28-29 页少了 12 个连续词条

accident、by accident、expect、silver、lining、silver lining、situation、care about、reach、reach for、touch、lend (sb) a hand。
CSV 从 exchange / Rose 直接跳到了 Unit 4，这一整列被丢掉。
另外 emotional intelligence（p.66）也缺失。

### 2. 页码小错位（多为差 1）

get together p.7→p.8；countryside p.9→p.8；Scotland p.9→p.7；slim p.25→p.24；fact p.25→p.24；sense p.27→p.26。

### 3. 释义或音标小差异

scarf 图片"围巾；披巾" → CSV"围巾；头巾"。
prize 图片"奖；奖励" → CSV"奖；奖品"。
pleasant 图片"宜人的" → CSV"令人愉快的"。
ancient 音标图片 /ˈeɪnʃənt/ → CSV /ˈenʃənt/（少一个 ɪ）。

### 4. 非错误，仅体例不同

km / kg / mm / per cent：CSV 把 "(= kilometre)" 等注释并进了单词字段。
spoon 拆成 spoon + spoonful 两条；yourself 拆成 yourself + yourselves 两条。
8A 音标整体用方括号且没有 7A 那种 l→I 的系统损坏。

---

## 7B / 8B / 9A / 9B —— 一致，可放心

7B：图片 461，CSV 461，遗漏 0、多余 0。仅 5 处微差（kick 图片"踹" vs CSV"蹬"；museum、hopefully、feather 音标少个别符号）。

8B：图片 601，CSV 601，遗漏 0、多余 0。仅 11 处音标记号级微差。

9A：图片 486，CSV 485。仅漏 1 词 widely（Unit 5 · p.48）。其余为音标记号级微差。

9B：图片 247，CSV 247。遗漏/多余各 2 条实为同词括号写法差异，非真错。字段差异多为低清页我方转写近似，CSV 忠实源自 PNG。

---

## 建议

1. 7A：从 7a word-list PNG 重新生成，重点修复音标 l→I 的系统性损坏，并补齐 19 个遗漏词、清除 4 个多余词、校正页码。
2. 8A：补上 Unit 3 第 28-29 页缺失的 12 个连续词条与 emotional intelligence。
3. 9A：补上 widely 一词。
4. 7B / 8B / 9B：无需改动。

---

## 修复记录（2026-07-13 已执行）

已备份原文件为 merged_vocabulary.backup-2026-07-13.csv。

1. 7A：整册按 7a word-list PNG 重建（318 → 333 行）。
   修复音标 l→I 系统性损坏；补齐 19 个遗漏词；删除 4 个多余词（含 CD）；
   页码与单元名改回 PNG（如 each 回到 Starter Unit 1 · p.2）。
   该册 source_file 改为 7a-word-list_01-07.png，sequence 置空，页码改为 p.N 体例。
2. 8A：在 exchange 之后补入 Unit 3 缺失的 12 个连续词条；在 intelligence 之后补入 emotional intelligence（540 → 553 行）。
3. 9A：在 gradually 之后补入 widely（485 → 486 行）。

复验结果：7A 与 9A 与 PNG 完全对齐（遗漏/多余/字段差异均为 0，9A 仅剩音标记号级微差）。
8A 真实缺失已全部补齐；剩余 6 条差异为词条边界/注释体例不同（如 km(= kilometre)、mashed potatoes），非错误，保留原体例。

总行数 2652 → 2681。

说明：重建后的 7A 内容来自对 PNG 的视觉转写，虽远优于原损坏数据，仍建议抽查复核。
