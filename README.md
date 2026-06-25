# GRE-bmw Excel Skill

A Trae skill that extracts **word roots, prefixes, and suffixes** from user-uploaded **handwritten GRE note screenshots** and appends them to the existing `cctalk GRE BMW.xlsx` notebook.

## What it does

1. Reads a handwritten GRE note screenshot uploaded by the user
2. The user runs OCR (e.g. via WeChat / Feishu / 夸克 / 白描) and pastes the recognized text back
3. The skill parses the text into structured entries (term / meaning / examples / mnemonic device)
4. **For each image, the user confirms the OCR result before the skill writes anything to Excel**
5. New entries are **appended** to the appropriate sheet (`prefix` / `root` / `suffix`), preserving existing data, headers, and titles
6. The `mythology&history` sheet is **not** processed — myth/historical content in screenshots is ignored

## Target file

A single, fixed Excel file:

```
cctalk GRE BMW.xlsx
```

The skill never creates a new file, never modifies the title row (row 1) or the header row (row 2), and never deletes existing data. It only appends rows starting from row 3, with a new auto-incremented number in column A.

## Sheet structure (must match exactly)

| Sheet | Title (row 1) | Header (row 2) | Columns |
|-------|---------------|----------------|---------|
| `prefix` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Prefix` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Prefix / C Meaning / D Examples / E Mnemotic |
| `root` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Root` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Root / C Meaning / D Examples / E Mnemotic |
| `suffix` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Suffix` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Suffix / C Meaning / D Examples / E Mnemotic |
| `mythology&history` | `神话与历史词源` | `词源` / `典故` / `重点词汇` | *(ignored by this skill)* |

## When to invoke this skill

The skill should be activated when the user:

- Uploads a handwritten GRE vocabulary screenshot
- Asks to "整理到 Excel" / "写入笔记" / "补全笔记" / "提取词根前缀后缀"
- Mentions "接着上一个笔记写" / "继续填这个表"

## Workflow

1. **Read image** — user provides the OCR text (this environment has no direct OCR capability)
2. **Confirm per image** — paste the OCR text back to the user and wait for confirmation before writing
3. **Parse** — extract term / meaning / examples / mnemonic device from each entry
4. **Route to sheet** — prefix/root/suffix only; myth/historical content is skipped
5. **Append to Excel** — write to the next available row with an auto-incremented number
6. **Report** — list how many entries were written to which sheet, and which rows they landed on

## Multi-image rule

When the user uploads multiple images:

- The user pastes OCR text for **image 1 first**, the skill writes it, **then** image 2, then 3, ...
- The user confirms the OCR result **once per image** (never all at once)
- Within an image, if entries belong to multiple sheets, they are written in the order they appear in the OCR text

## Examples field format

The `Examples` column keeps the same separator as the source image:

- If the image shows `dangerous/arduous`, write `dangerous/arduous`
- If the image shows `dangerous, arduous`, write `dangerous, arduous`
- If the image shows them on separate lines, join them with `/`

## Mnemonic device

The `Mnemotic Device` column is **actively inspected** for any mnemonic / etymology / cultural reference in the image. If present, it is written. If not, it is left blank — never invented.

## Limitations

- This skill **cannot read images directly**. The user must run external OCR and paste the result.
- This skill **cannot** process the `mythology&history` sheet.
- This skill is **append-only**. It does not sort, dedupe, or modify existing rows.

## Installation

This is a [Trae](https://trae.ai) skill. To use it:

1. Place the `SKILL.md` file under `.trae/skills/GRE-bmw/` in your Trae workspace
2. Place your `cctalk GRE BMW.xlsx` file at the location referenced inside `SKILL.md`
3. Trae will auto-discover the skill based on the `name` and `description` in the frontmatter

## Files in this repository

- `SKILL.md` — the skill definition (frontmatter + body)
- `README.md` — this file
- `LICENSE` — MIT License
