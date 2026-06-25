---
name: "GRE-bmw"
description: "Extract word roots, prefixes, and suffixes from user-uploaded handwritten GRE note screenshots, confirm OCR text with the user per image, and append them to the existing cctalk GRE BMW.xlsx notebook (prefix/root/suffix sheets only). Invoke when user uploads a handwritten vocabulary screenshot and asks to 整理词根前缀后缀/写入Excel/整理到笔记/补全笔记."
---

# GRE-bmw 技能

## 技能目的

从用户上传的 **手写 GRE 笔记截图** 中识别**词根、前缀、后缀**等构词要素，并按用户的现有笔记格式，**追加/补全** 到现有的 `cctalk GRE BMW.xlsx` 笔记文件中，节省抄写时间。

**注意：本技能只处理 prefix / root / suffix 三类构词要素，不处理 mythology&history（神话历史）内容。**

## 触发场景

**当用户出现以下意图时调用本技能：**

- 上传一张或一组 **手写 GRE 笔记截图**，并要求"整理到 Excel"、"写入笔记"
- 提到"提取词根"、"补全笔记"、"追加到现有的 Excel"
- 类似的关键词：**词根 / 前缀 / 后缀 / 词缀 / 整理 / 补全 / 笔记 / cctalk / BMW**
- 用户表达"接着上一个笔记写"、"继续填这个表"等需求

**典型用户表述：**
- "把这张手写笔记的词根整理到 cctalk GRE BMW.xlsx 里"
- "接着把 prefix 写完"
- "补全 root 那张表"

## 目标文件

**固定使用**：`d:\桌面\项目\skills\GRE单词\cctalk GRE BMW.xlsx`

## Excel 结构（严格遵循）

文件包含 4 个 sheet，结构如下：

| Sheet 名 | 第 1 行（标题） | 第 2 行（表头） | 数据列 |
|----------|----------------|----------------|--------|
| `prefix` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Prefix` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Prefix / C Meaning / D Examples / E Mnemotic |
| `root` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Root` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Root / C Meaning / D Examples / E Mnemotic |
| `suffix` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Suffix` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Suffix / C Meaning / D Examples / E Mnemotic |
| `mythology&history` | `神话与历史词源` | `词源` / `典故` / `重点词汇` | A 编号 / B 词源 / C 典故 / D 重点词汇 |

**重要规则：**
- 不要新增 sheet、不要改表头、不要改标题行
- 数据从第 3 行开始，A 列填递增编号（1, 2, 3...）
- 编号需要从**该 sheet 当前最大编号 + 1** 开始追加
- 第 2 行是表头，**不要覆盖**

## 工作流程

### Step 1：读取图片（借助外部 OCR）

由于本环境**不具备直接 OCR 能力**，识别流程如下：

- 用户上传手写笔记截图后，先将图片提交到外部 OCR 服务（或由用户提供 OCR 后的文字版本）
- 当前推荐的外部 OCR 方案：
  - 飞书/微信自带的图片文字提取
  - 手机端 QQ 浏览器 / 夸克 / 百度网盘等的"提取文字"功能
  - 专业 OCR 工具（白描、夸克扫描王等）
- 一次处理用户提供的所有图片，**不漏读**
- 注意手写可能不规范，**识别不确定的标记为待确认**

**前置确认（强制流程，每张图都做）：**

每张图片的 OCR 结果都必须在写入 Excel 之前与用户确认：

1. **先把 OCR 文字贴给用户**："图 1 OCR 结果如下：[贴出]。请确认是否准确？有错请直接改正。"
2. **等待用户确认或修改**
3. 用户说"确认 / 没问题 / OK / 写入" 等之后再进入 Step 2 解析该张图片
4. 下一张图重复此流程

**确认前的状态：**
- 暂不写入 Excel
- 不进入 Step 2、3、4
- 如果用户一次性上传多张图，按图片顺序**逐张**确认（不要一次性让用户确认所有图）

### Step 2：识别每条词条

从图片中识别每条词条，按以下字段提取：

| 字段 | 说明 | 示例 |
|------|------|------|
| **类型** | prefix / root / suffix / 神话历史 | `prefix` |
| **词缀/词根/词源** | 主体内容 | `ab-, abs-` |
| **含义** | 释义（英文为主，与原笔记一致） | `away from, off` |
| **举例 (Examples)** | 派生词 | `dangerous/arduous` |
| **助记/典故 (Mnemotic Device)** | **图片中有的助记法/典故也要写入** | （有就写，没有就留空） |

**Examples 字段格式说明：**

- 不强制使用 `/` 或 `,` 分隔，**保持图片中原始分隔方式**
- 如果原图用的是 `/`，就用 `/`
- 如果原图用的是 `,`，就用 `,`
- 如果原图是换行，就用 `/` 合并
- 最终目的：与图片中的视觉风格保持一致

**Mnemotic Device 字段说明：**

- **必须主动检查图片中是否包含助记法或典故**
- 如果图片里有"谐音"、"形象记忆"、"希腊神话 XXX"、"罗马 XXX"等字样，必须提取并写入
- 图片中无相关内容则留空（不要瞎编）

### Step 3：判断写入哪个 Sheet（按用户提供的图片顺序逐个 sheet 追加）

**追加规则：按用户上传图片的顺序，逐个 sheet 追加，不打乱顺序。**

- 第 1 张图片 → 解析后写入对应的 sheet（prefix/root/suffix）
- 第 2 张图片 → 接着上一张继续写同一个 sheet 或新 sheet（按图片内容判断）
- 第 3, 4, 5... 张图片 → 依此类推

**关于 mythology&history sheet：**

- **本技能不处理 mythology&history sheet**
- 如果图片中包含神话/历史典故内容，**忽略这部分**，不写入任何 sheet
- 仅处理 prefix / root / suffix 三类构词要素

**判断类型：**
- 前缀（ab-, in-, un-...）→ `prefix`
- 词根（bene, dict, ject...）→ `root`
- 后缀（-ous, -able, -itive...）→ `suffix`
- 神话/历史典故 → **忽略，不写入**

### Step 4：追加写入 Excel

使用 `openpyxl`（或 `xlsx` skill）追加到目标文件：

```python
import openpyxl
wb = openpyxl.load_workbook('d:/桌面/项目/skills/GRE单词/cctalk GRE BMW.xlsx')
ws = wb['prefix']  # 或 root / suffix / mythology&history

# 找到下一行（从第 3 行开始，找 A 列最大编号）
max_row = ws.max_row
start_num = 1
for row in range(3, max_row + 1):
    val = ws.cell(row=row, column=1).value
    if isinstance(val, int) and val >= start_num:
        start_num = val + 1

# 写入新数据
for i, item in enumerate(extracted_items):
    r = max_row + 1 + i  # 或基于 start_num 计算行号
    ws.cell(row=r, column=1, value=start_num + i)  # 编号
    ws.cell(row=r, column=2, value=item['term'])
    ws.cell(row=r, column=3, value=item['meaning'])
    ws.cell(row=r, column=4, value=item['examples'])
    ws.cell(row=r, column=5, value=item.get('mnemonic', ''))

wb.save('d:/桌面/项目/skills/GRE单词/cctalk GRE BMW.xlsx')
```

**写入规则：**
- A 列编号：从该 sheet 现有最大编号 + 1 开始
- 跳过空行
- **不修改** 第 1 行标题、第 2 行表头
- **不删除** 已存在的数据
- 保持原有格式（字体、对齐等）

### Step 5：反馈结果

向用户汇报：

1. 共识别了多少条词条
2. 分别写入到哪个 sheet（prefix/root/suffix）
3. 各 sheet 当前的总条目数
4. **明确的文件路径**：`d:\桌面\项目\skills\GRE单词\cctalk GRE BMW.xlsx`
5. 对于识别不确定的词条，**单独列出**让用户确认
6. **说明跳过的神话/历史内容数量**（如果有）

## 输出格式

每次调用结束后，输出必须包含：

```
✅ 识别完成
- 共处理图片：N 张
- 写入顺序（按图片顺序）：
  1. prefix:  +3 条（图片 1）→ 第 21-23 行
  2. root:    +5 条（图片 2）→ 第 50-54 行
  3. prefix:  +4 条（图片 3）→ 第 24-27 行
  4. suffix:  +2 条（图片 3）→ 第 5-6 行
  5. root:    +6 条（图片 4）→ 第 55-60 行
- 跳过内容：图片 5 含 3 条神话/历史（按规则不写入）
- 文件路径：d:\桌面\项目\skills\GRE单词\cctalk GRE BMW.xlsx
- 待确认词条（如有）：[列出]
```

## 注意事项

- **绝不要覆盖** 已存在的数据，只追加
- **绝不要修改** 第 1 行标题、第 2 行表头
- 手写体识别可能不准确，**主动询问**而不是猜测
- 保持原笔记的英文风格（不要把 `away from` 翻译成中文）
- 处理多张图片时，**严格按照用户上传顺序逐个 sheet 追加**
- 多次调用本技能时，编号会自然递增，不需要重置
- Mnemotic Device 字段**必须主动检查**图片中是否有助记法/典故，有就写，没有就留空
- Examples 字段保持图片中的原始分隔方式（`/` 或 `,`）
- **每张图的 OCR 结果都必须经用户确认后再写入**，不允许跳过
- **不处理 mythology&history sheet**，图片中遇到神话/历史内容直接忽略

## 当前 Excel 状态（截至初始化）

- `prefix`: 已有数据 277 行（含标题/表头），实际词条从第 3 行起
- `root`: 已有数据 276 行
- `suffix`: 已有数据 28 行
- `mythology&history`: 已有数据 50 行

（注：以上为初始状态，调用过程中会变化）
