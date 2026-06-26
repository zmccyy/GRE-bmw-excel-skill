---
name: "GRE-bmw"
description: "Extract word roots, prefixes, suffixes, and mythology/history allusions from user-uploaded handwritten GRE note screenshots, confirm OCR text with the user per image, and append them to the existing cctalk GRE BMW.xlsx notebook. Use TWO modes: (A) root/prefix/suffix mode writes to prefix/root/suffix sheets when right side has red-dot rows; (B) mythology mode writes to mythology&history sheet when right side has NO red-dot row and there's a black bold proper noun (e.g. Olympus, Bacchus) followed by an English story sentence. Invoke when user uploads a handwritten vocabulary screenshot and asks to 整理词根前缀后缀/写入Excel/整理到笔记/补全笔记/添加神话典故/神话与历史."
---

# GRE-bmw 技能

## 技能目的

从用户上传的 **手写 GRE 笔记截图** 中识别**词根、前缀、后缀**等构词要素，以及**神话/历史典故**（Olympus / Bacchus 等），按用户的现有笔记格式，**追加/补全** 到现有的 `cctalk GRE BMW.xlsx` 笔记文件中，节省抄写时间。

**技能支持两种模式**：
1. **词根/前缀模式**（默认）：处理 prefix / root / suffix 三类构词要素
2. **神话典故模式**（特殊）：处理 mythology&history 类的词源典故（详细规则见 Step 2.7）

## 触发场景

**当用户出现以下意图时调用本技能：**

- 上传一张或一组 **手写 GRE 笔记截图**，并要求"整理到 Excel"、"写入笔记"
- 提到"提取词根"、"补全笔记"、"追加到现有的 Excel"
- 上传**神话典故类截图**（如 Olympus / Bacchus / Achilles 等），要求"写入神话"、"添加神话"
- 类似的关键词：**词根 / 前缀 / 后缀 / 词缀 / 整理 / 补全 / 笔记 / cctalk / BMW / 神话 / 典故 / 词源 / Olympus / Bacchus / Achilles / Pandora**
- 用户表达"接着上一个笔记写"、"继续填这个表"等需求

**典型用户表述：**
- "把这张手写笔记的词根整理到 cctalk GRE BMW.xlsx 里"
- "接着把 prefix 写完"
- "补全 root 那张表"
- "把 Bacchus 这个神话典故加到 mythology&history sheet"

## 目标文件

**固定使用**：`d:\桌面\项目\skills\GRE单词\cctalk GRE BMW.xlsx`

## Excel 结构（严格遵循）

文件包含 4 个 sheet，结构如下：

| Sheet 名 | 第 1 行（标题） | 第 2 行（表头） | 数据列 |
|----------|----------------|----------------|--------|
| `prefix` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Prefix` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Prefix / C Meaning / D Examples / E Mnemotic |
| `root` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Root` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Root / C Meaning / D Examples / E Mnemotic |
| `suffix` | `cctalk GRE BMW词根缀课程-配套笔记模板` | `Suffix` / `Meaning` / `Examples` / `Mnemotic Device` | A 编号 / B Suffix / C Meaning / D Examples / E Mnemotic |
| `mythology&history` | `神话与历史词源`（A1:D1 合并）| `编号` / `词源` / `典故` / `重点词汇` | A 编号 / B 词源 / C 典故 / D 重点词汇 |

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


### Step 2：标准化布局解析（图片布局规则）

**⚠️ 这一步是核心解析逻辑，必须严格按以下布局规则执行。**

用户的 GRE 笔记截图采用**标准化双栏布局**，必须按视觉结构逐区解析：

#### 2.1 整体画布划分

将图片大致**垂直分为左右两大功能区**：

| 区域 | 位置 | 内容 | 字段 |
|------|------|------|------|
| **左侧区域** | 左侧 | 红色方框（或红色大括号）框住的部分 | **Mnemotic Device（助记法/语境）** |
| **右侧区域** | 右侧 | 词根词缀核心数据 | 词根/词缀、含义、举例 |

#### 2.1.1 图片模式判定（必须先做！）

**重要**：在解析前必须先判断这张图属于哪种模式，**不同模式走不同解析路径**。

| 模式 | 触发条件 | 目标 sheet |
|------|----------|-----------|
| **A. 词根/前缀模式** | 右侧有 **红点行（■/●/•）** + 蓝/红色高亮词根 | `prefix` / `root` / `suffix` |
| **B. 神话典故模式** | 右侧**无红点行** + 左侧出现**单个黑体大写专有名词**（如 Olympus）+ 下方紧跟完整英文故事句 | `mythology&history` |

**判定优先级**：B 优先于 A（即优先检查是否符合神话典故模式）

**判定流程**：
1. 扫描右侧区域，**确认是否有红点行**（`■` `●` `•` 开头）
2. 如果**没有红点行** → 进一步检查左侧是否出现"黑体大写专有名词 + 完整故事句"模式
3. 如果符合 → 启用 **2.7 神话典故模式**解析
4. 如果不符合 → 退回 **2.3 词根/前缀模式**解析

#### 2.2 左侧区域解析（Mnemotic Device）

- 由**红色方框**或**红色大括号**框出
- 内含若干条**完整的英文句子或短语搭配**
- 全部内容提取为 `Mnemotic Device` 字段
- 保留原文的标点和换行

#### 2.3 右侧区域解析（核心数据）

**严格按照换行和颜色执行**：

**第 1 行（词根/前缀 & 含义）：**

- 寻找带 **红点（■/●/•）标记** 的行（即 `■` `●` `•` 这类小圆点符号开头的那一行）
- 红点行的内容是**核心词根或前缀**，**不是**例词
- 同一行内可能既有前缀也有词根（如 `pre` + `diction`），它们之间用空格或连字符分隔
- **类型判断规则（按颜色，与红点无关）：**
  - **红色**词 → 归类到 **prefix**（前缀）
  - **蓝色**词 → 归类到 **root**（词根）
  - 其它颜色 → 按常见分类判断（-ous/-able/-itive → suffix）
- **含义提取：**
  - 在同一行内，紧挨着核心词（通常是逗号或空格后）
  - 跟着的一串简短英文（如 `to say` 或 `before`）
  - 提取为 `Meaning`（含义）

**红点行下方（例词列表）：**

- 在红点行下方，**所有换行罗列的单词**（如 `prediction`、`indicate`、`contradict` 等）
- **全部都是例词**（不区分颜色，红色蓝色黑色一律视为例词）
- 全部抓取并合并为 `Examples`（例词）列
- **用 `,` 隔开**（与图片中的换行视觉对齐）

#### 2.4 特殊情况处理

**变体处理（仅指同一类型的多个写法，如 `ben, bon` 或 `pre/ante/fore`）：**

- 一行有多个变体（如 `ben, bon` 或 `pre/ante/fore`）：
  - **用 `,` 分割的** → 放在**同一格**
  - **用 `/` 分割的** → 每个**单独一格**（每个变体各自生成一条记录）
- 例子：
  - 输入 `ben, bon` + 含义 `good, well` → **1 条记录**，`term="ben, bon"`, `meaning="good, well"`
  - 输入 `pre/ante/fore` + 含义 `before` → **3 条记录**，每条 `term=pre/ante/fore` 中的一个

**红点行内的"前缀+词根"组合（如 `pre` + `diction`）：**

- **不拆分成多条记录**
- 把整行视为**一个完整的构词单元**
- 写入时只写 **1 条**记录，词缀列写**整个组合**（如 `pre-diction` 或 `pre diction`，与图片保持一致）
- 类型归属按**主词根**的颜色判断（`diction` 蓝色 → 归 `root`）

**单张图含 2-3 组词根/前缀（多组情况）：**

一张图中可能出现 **2-3 个独立的红点行**，每行代表一组独立的词根/前缀。这种情况**完全正常**，按以下规则处理：

- 每个红点行 = 一组独立的词根/前缀数据
- 每组之间用**空行**或**明显的视觉分隔**区分（如大字距、缩进变化）
- 每组分别独立解析（按 2.3 的红点行规则）
- 每组分别匹配左侧的助记句（按左↔右垂直对齐）

**多组识别流程：**

1. **扫描整个右侧区域**，找到所有带红点（■/●/•）的行
2. **确定组数 N**（N 通常为 1-3）
3. **从每个红点行往下扫描**，收集该组的例词列表，直到：
   - 遇到下一个红点行（该行属于下一组）
   - 遇到空行 + 新红点行
   - 到达右侧区域末尾
4. **每组独立生成 1 条或多条记录**（根据 `,` 或 `/` 分割规则）

**典型多组示例：**

```
[右侧]
■ ab, abs    away, from
abdicate
abduct
abscond
abstemious
abstract
abstruse

■ pre/ante/fore    before
predict
prepare
preference
...

■ contra, contro/anti    against
contradictory
controversial
...
```

这是 **3 组独立的核心词根/前缀**，每组分别生成自己的记录（ab, abs → 1 条；pre/ante/fore → 3 条；contra, contro/anti → 2 条），共 6 条记录。

**多组与左侧助记的匹配规则：**

- **一对一匹配（最常见）**：左侧第 1 句 → 右侧第 1 组；左侧第 2 句 → 右侧第 2 组；...
- **一对多匹配**：左侧 1 句 → 右侧 N 组（N 个红点行在视觉上"共属"一个主题）—— 这种情况下该助记句**复制**到所有 N 组的 Mnemotic 字段
- **多对一匹配**：左侧 N 句 → 右侧 1 组 —— N 句**全部合并**到该组的 Mnemotic 字段（用 `;` 分隔）
- **完全错位**：M 句 ↔ N 组无法对应 → 单独列出让用户确认

**多组写入时的关键点：**

- 写入顺序按图片中**从上到下**的顺序
- 每组都从**该 sheet 当前最大编号 + 1** 开始
- 编号连续递增（不跳号）
- 不需要分批写入，一次 openpyxl 调用完成

**左↔右垂直对齐匹配：**

- 左侧的句子与右侧的词组按**视觉垂直对齐**进行匹配
- 匹配规则：
  - 左侧第 1 句 → 右侧第 1 组词根（最右栏）
  - 左侧第 2 句 → 右侧第 2 组词根
  - ...
- 如果左侧有 N 句、右侧有 M 组（M ≠ N）：
  - **优先用原词匹配**：如果左侧句子中包含右侧某个具体变体（如左侧出现 "pre-"，右侧有 "pre/ante/fore"），将这一句绑定到该变体
  - 多余的句子作为该词根的**通用助记法**附在第一条记录
  - 缺失的句子位置留空
- 如果完全无法匹配，**单独列出**让用户确认

#### 2.5 解析示例

**OCR 原始文本（一张典型笔记图，含红点行）：**

```
[左侧红色方框]                          [右侧]
                                        ■ dic, dict    to say       <- 红点行，蓝色=root
The time before the meeting
  was tense.                            prediction
                                        indicate
The soldiers advanced toward
  the front line.                       contradict
                                        abdicate
                                        diction
                                        dictum
```

**关键识别点：**
- `■ dic, dict` 这行有红点 `■`，是**核心词根行**（不是例词）
- `dic, dict` 是**蓝色**，按规则归类为 **root**
- `to say` 在同一行内 → **含义**
- 下方 6 个词（`prediction`/`indicate`/`contradict`/`abdicate`/`diction`/`dictum`）→ 全部是**例词**，不区分颜色

**解析后（写入 root sheet 的 1 条记录）：**

| 类型 | 词根/词缀 | 含义 | Examples | Mnemotic |
|------|----------|------|----------|----------|
| root | `dic, dict` | to say | prediction, indicate, contradict, abdicate, diction, dictum | The time before the meeting was tense. The soldiers advanced toward the front line. |

（注：`dic, dict` 用 `,` 分割 → 1 条记录；下方所有例词不区分颜色一律合并到 Examples）

#### 2.6 字段填写规则

| 字段 | 来源 | 示例 |
|------|------|------|
| **类型** | 颜色判定（红=prefix, 蓝=root） | `prefix` |
| **词根/词缀** | 第 1 行高亮词组 | `pre` 或 `ben, bon` |
| **含义** | 第 1 行高亮词后的英文 | `before` 或 `good, well` |
| **Examples** | 第 3 行及以下的换行列出的单词 | `predict, pretest, preview` |
| **Mnemotic Device** | 左侧红色方框中的句子 | `The time before the meeting was tense.` |

**Examples 字段格式说明：**

- 右侧第 3 行及以下的每个单词用 `,` 分隔（**不区分**图片中原本是 `/` 还是 `,` 还是换行）
- 目的：与图片中"每个单词一行"的视觉风格保持一致

**Mnemotic Device 字段说明：**

- **必须主动检查**图片左侧是否有红色方框/大括号
- 左侧句子直接整段提取（保留标点、保留换行或用空格连接均可）
- 如果左侧无红色方框/大括号 → 留空（**不要瞎编**）

#### 2.7 神话典故模式（Mythology Mode）

**适用范围**：经过 2.1.1 模式判定，符合"神话典故模式"的图片。

**重要更新**：本技能**现在也处理 mythology&history sheet**（与之前"不处理"的规则不同）。当图片被判定为神话典故模式时，**写入 `mythology&history` sheet**。

##### 2.7.1 触发条件（同时满足）

1. 图片右侧**没有红点行**（`■` `●` `•` 开头）
2. 图片中出现**单个黑体大写专有名词**（如 `Olympus`、`Achilles`、`Pandora`），通常：
   - 全大写或首字母大写
   - 黑体/加粗显示
   - 紧跟在**红色方框或红色横条**后方
3. 该黑体词**正下方**紧跟一句**完整的英文故事句**（作为记忆锚点）

##### 2.7.2 提取映射规则（严格按视觉层次）

| 列 | 视觉位置 | 字体特征 | 示例 | 写入字段 |
|----|----------|----------|------|----------|
| **第 1 列** | 红色方框（或红条）后方紧跟 | 黑色粗体单词 | `Olympus` | `词源/名称`（B 列） |
| **第 2 列** | 该黑体词正下方、缩进或换行 | 浅色/常规字重的完整英文句子 | `The Greek Gods lived high atop Mount Olympus.` | `背景故事/语境`（C 列） |
| **第 3 列** | 段落末尾 | 紫色或特殊颜色标注的词性及中文释义 | `olympian a. 高傲的，疏离的` | `派生词汇`（D 列） |

**关键处理细节**：

- **第 1 列（词源/名称）**：
  - 提取红色方框后方紧跟的**黑色粗体单词**
  - 通常是神话人物、神话地点、神祇名（如 `Olympus`、`Achilles`、`Pandora`）
  - 保留原始大小写（专有名词首字母大写或全大写）

- **第 2 列（神话语境/出处）**：
  - 提取该黑体词**正下方**、**缩进或换行显示**的完整英文句子
  - **浅色或常规字重**（区别于黑体粗体）
  - 作为**记忆锚点**
  - 保留完整句子（保留标点、句号）
  - 通常是 1 句话，也可能 2-3 句

- **第 3 列（衍生词/拓展词汇）**：
  - 提取该段落**末尾**出现的、用**紫色或特殊颜色**标注的内容
  - 格式通常是：`词 词性. 中文释义`（如 `olympian a. 高傲的，疏离的`）
  - **必须去掉前面的短横线或项目符号**（`-` `•` `·` 等）
  - 保留原始的**词性符号**（`a.` `n.` `v.` `ad.` 等）
  - 如果有多个衍生词，**用 `;` 分隔**（每个都按 "词 词性. 中文释义" 格式）

##### 2.7.3 表格填充目标

写入 **`mythology&history` sheet`，4 列严格对应**：

| mythology&history 列 | Excel 实际位置 | 内容来源 | 用户指定字段名 |
|---------------------|---------------|----------|----------------|
| A 编号 | A 列 | 自动递增（1, 2, 3...） | **编号** |
| B 词源 | B 列 | 第 1 列（红色方框后黑体词，如 `Bacchus`） | **词源/名称** |
| C 典故 | C 列 | 第 2 列（黑体词下方完整英文句子） | **背景故事/语境** |
| D 重点词汇 | D 列 | 第 3 列（紫色字体衍生词） | **派生词汇** |

**实测表头**（写入前自动验证）：
- 第 1 行：`神话与历史词源`（A1:D1 合并单元格）
- 第 2 行：A=`编号`, B=`词源`, C=`典故`, D=`重点词汇`

**填充示例**：

| A 列（编号） | B 列（词源） | C 列（典故） | D 列（重点词汇） |
|--------------|--------------|--------------|------------------|
| `48` | `Bacchus` | `Roman god of drama, wine, and ecstasy` | `Bacchanalian a. 饮酒作乐的` |

##### 2.7.4 解析示例

**OCR 原始文本（一张典型神话典故图）：**

```
[左侧红色横条]                            [右侧]
[Olympus] 黑体粗体                       
The Greek Gods lived high atop           （右侧无红点行）
  Mount Olympus.                         
                                         olympian a. 高傲的，疏离的   <- 紫色字体
```

**关键识别点**：
- 右侧**无红点行** → 触发模式 B
- `Olympus` 是**黑体粗体** + 紧跟**红色横条** → 词源
- 下方缩进完整句子 `The Greek Gods lived high atop Mount Olympus.` → 典故
- 末尾紫色字体 `olympian a. 高傲的，疏离的` → 派生词

**解析后（写入 mythology&history sheet 的 1 条记录）**：

| 编号 (A) | 词源 (B) | 典故 (C) | 重点词汇 (D) |
|----------|----------|----------|--------------|
| `49` | `Olympus` | `The Greek Gods lived high atop Mount Olympus.` | `olympian a. 高傲的，疏离的` |

##### 2.7.5 与词根/前缀模式的区别

| 维度 | 词根/前缀模式（2.3） | 神话典故模式（2.7） |
|------|---------------------|---------------------|
| 目标 sheet | `prefix` / `root` / `suffix` | `mythology&history` |
| 右侧标识 | 有红点行 | 无红点行 |
| 左侧内容 | 红色方框包住的英文短语/句子 | 红条/红方框 + 黑体大写专有名词 |
| 字段数 | 5 列（含 Mnemotic Device） | 4 列（无 Mnemotic Device） |
| 颜色含义 | 红=prefix, 蓝=root | 仅区分黑体（词源）/ 浅色（典故）/ 紫色（派生） |
| 例子数 | 多行例词列表 | 1-2 个衍生词 |

##### 2.7.6 特殊情况

- **多张图都是神话典故**：每张图生成 1 条记录，按图片顺序追加到 mythology&history sheet
- **同一张图多个神祇/典故**：按图片中**从上到下**顺序解析，每个生成 1 条记录
- **图 1 是词根模式 + 图 2 是神话模式**：**先写词根模式的 sheet，再写 mythology&history sheet**，按用户上传顺序追加

##### 2.7.7 实战示例：Bacchus（用户笔记原图）

**OCR 原始文本（用户上传的 Bacchus 神话图）：**

```
[左侧红色方框 - Mnemonics]              [右侧 - 背景说明]
@ Sherry Zhang                          酒神是古希腊及古罗马...

■ Bacchus                               ...（右侧内容是中文背景介绍，
Roman god of drama, wine, and ecstasy     不是结构化数据，应忽略）
Bacchanalian a. 饮酒作乐的
```

**关键识别点**：
- 右侧**无红点行** + 左侧有 `Bacchus` 黑体粗体 → 触发**模式 B（神话典故）**
- 右侧内容是中文背景介绍，**忽略**（只关心左侧红色方框内的内容）
- 红色方框内有 3 个层级：
  1. `■ Bacchus`（红色 `■` 标记 + 黑体大写专有名词）→ **词源/名称**
  2. `Roman god of drama, wine, and ecstasy`（浅色常规字重完整英文句子）→ **典故/背景故事**
  3. `Bacchanalian a. 饮酒作乐的`（**紫色字体**）→ **派生词汇**

**注意修正**：
- 用户笔记里词源前的 `■` 是红色方块标记（不是颜色判定），但因为词源本身是 `Bacchus`（专有名词，黑色加粗），所以仍归神话典故模式
- 与通用"红点行=核心词根"的规则区别：这里的 `■` 后跟的是**专有名词**（不是泛词根），且下面**没有例词列表**

**解析后（写入 mythology&history sheet 的 1 条记录）**：

| 编号 (A) | 词源 (B) | 典故 (C) | 重点词汇 (D) |
|----------|----------|----------|--------------|
| `48` | `Bacchus` | `Roman god of drama, wine, and ecstasy` | `Bacchanalian a. 饮酒作乐的` |

##### 2.7.8 神话典故模式的"实操判定捷径"

由于神话典故图片的视觉特征独特，可以用以下**快速判定捷径**：

| 捷径 | 检查 | 通过条件 |
|------|------|----------|
| 1️⃣ 右侧有红点行？ | 扫描 `■` `●` `•` | **没有**才能走模式 B |
| 2️⃣ 左侧有黑体大写专有名词？ | 全大写 OR 首字母大写 + 黑体/加粗 | **是** |
| 3️⃣ 该专有名词下方有完整英文句子？ | 普通字重 + 句号结尾 | **是** |
| 4️⃣ 段落末尾有紫色字体衍生词？ | 紫色 + 词性符号 + 中文释义 | 通常有（但非必须） |

**3 个捷径都通过 → 确认是神话典故模式**，启用 2.7 解析。

**如果只有 1-2 捷径通过** → 单独询问用户确认模式。

##### 2.7.9 列结构修正记录（2026-06-26）

**之前的填充方式（错误）**：A 列放"编号 + 词源"的合并字符串（如 `48 Bacchus`）
**用户最新要求（正确）**：A=编号, B=词源, C=典故, D=重点词汇（4 列严格分开）

**已执行的修复**：
1. **Excel 表头修正**：
   - A2：`词源` → `编号`
   - B2：新增 `词源`
   - C2/D2：保持 `典故` / `重点词汇`
   - A1:D1 标题行重新合并 + 补"神话与历史词源"
2. **Bacchus 第 50 行重写**：
   - A=48, B=Bacchus, C=Roman god of drama, wine, and ecstasy, D=Bacchanalian a. 饮酒作乐的

**新结构确认**（写入前自动验证）：

| A 列 | B 列 | C 列 | D 列 |
|------|------|------|------|
| 编号（int）| 词源（str）| 典故（str）| 重点词汇（str） |

**示例代码**（写入 mythology 模式）：

```python
ws = wb['mythology&history']
new_row = last_data_row + 1
new_num = last_num + 1
ws.cell(row=new_row, column=1, value=new_num)  # A 列：编号
ws.cell(row=new_row, column=2, value=item['source'])  # B 列：词源
ws.cell(row=new_row, column=3, value=item['story'])  # C 列：典故
ws.cell(row=new_row, column=4, value=item['derived'])  # D 列：重点词汇
```

**保留向后兼容**：如果未来再次发现 mythology&history sheet 的表头不是 4 列结构（如 A 列又要复用），技能会**报错提示用户确认**，不会盲目写入。

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
- **mythology&history sheet 处理规则已更新**：现在支持处理神话典故模式（见 2.7）。判定方法：右侧无红点行 + 左侧有黑体大写专有名词 + 下方紧跟完整英文故事句

## 当前 Excel 状态（动态更新）

- `prefix`: 实际有数据的行 3-94，编号 1-92（含写入测试 5 的 1 条新增）
- `root`: 实际有数据的行 3-184，编号 1-182（含写入测试 5 的 8 条新增）
- `suffix`: 实际有数据的行 3-28，编号 1-26（截至初始化）
- `mythology&history`: 实际数据行 3-50，编号 1-48（含写入测试 4 的 1 条新增）

**注意：每次写入前必须重新探查 A 列最大编号和实际最后行，max_row 不可信（中间可能有大片空行）。**

**mythology&history sheet 表头结构（已读取确认，2026-06-26 修正）**：
- 第 1 行（标题）：`神话与历史词源`（A1:D1 合并）
- 第 2 行（表头）：A=`编号`, B=`词源`, C=`典故`, D=`重点词汇`
- 数据列位置：**A 列**（int 编号）, **B 列**（词源）, **C 列**（典故）, **D 列**（重点词汇）

## 写入测试记录

**测试 1（2026-06-25）**：
- 目标：root sheet
- 写入行：第 173 行
- 编号：171
- 内容：`dic, dict | to say | prediction, indicate, contradict, abdicate, diction, dictum | "We do not inherit the earth from our ancestors. We borrow it from our children." —Indian dictum`
- 结果：✅ 成功

**测试 2（2026-06-25）— 4 张图批量写入**：
- **图 1**（root sheet）：
  - 写入行：第 174 行
  - 编号：172
  - 内容：`ben, bon | good, well | benefit, beneficial, benefactor, beneficiary, benediction, benevolence, benign | a new year's benediction; enjoy a benign climate; a benevolent leader; benign tumor`
  - 结果：✅ 成功
- **图 2**（prefix sheet，`pre/ante/fore` 用 `/` 分割 → 3 条）：
  - 行 83 / 编号 81：`pre | before | predict, prepare, preference, precursor, pre-empt, preclude, precocious, predispose, prerequisite | pre-emptive arrest; preclude youths from drugs; precocious child from housework; predispose his children to be successful`
  - 行 84 / 编号 82：`ante | before | antediluvian, antecedent | diluvian Noah's ark`
  - 行 85 / 编号 83：`fore | before | forecast, foretell, forestall | forestall an energy crisis`
  - 结果：✅ 成功
- **图 3**（prefix sheet）：
  - 行 86 / 编号 84：`ab, abs | away, from | abdicate, abduct, abscond, abstemious, abstract, abstruse | abdicate=quit; "I wish I knew how to quit you."; abstemious diets`
  - 结果：✅ 成功
- **图 4**（prefix sheet，`contra, contro/anti` 混合分割 → 2 条）：
  - 行 87 / 编号 85：`contra, contro | against | contradictory, controversial, contravene, contraband, contrarian | contravene criminal laws`
  - 行 88 / 编号 86：`anti | against | antigen, antibiotics, antisocialist, antagonist, antithesis | （留空）`
  - 结果：✅ 成功
- **总计**：root +1，prefix +6，共 7 条
- **踩坑**：max_row 不可信，必须扫描找真实最后行；prefix max_row 报告 277 但实际数据只到行 82

**测试 3（2026-06-25）— 2 张多组图写入**：
- **图 1**（root sheet，2 组）：
  - 行 175 / 编号 173：`vol, val | wish | volunteer, malevolence, ambivalent | （留空）`（助记留空，因为左侧 4 句都跟 mal 相关）
  - 行 176 / 编号 174：`mal | bad | malevolent, malediction, malign v/a. =malicious, malnourished, malevolence | malign his ex-wife for cheating on him; counter malign influences; malice of his ex-wife; malicious attacks`
  - 结果：✅ 成功
- **图 2**（prefix sheet，`ambi/amphi/bi/di/dup` 用 `/` 分割 → 5 条）：
  - 行 89 / 编号 87：`ambi | two, both | ambivalence, ambiguous, ambidextrous | （留空）`
  - 行 90 / 编号 88：`amphi | two, both | （留空，图中无独立例词） | （留空）`
  - 行 91 / 编号 89：`bi | two, both | bilingual | （留空）`
  - 行 92 / 编号 90：`di | two, both | dilemma, divergence, dichotomy | verge; convergence`（verge / convergence 归到 di 组）
  - 行 93 / 编号 91：`dup | two, both | duplex, duplicity, duplicate | （留空）`
  - 结果：✅ 成功
- **总计**：root +2，prefix +5，共 7 条

**测试 4（2026-06-26）— 神话典故模式（Bacchus）**：
- **图**（mythology&history sheet，模式 B 触发）：
  - **判定依据**：右侧无红点行 + 左侧有黑体大写专有名词 `Bacchus` + 下方有完整英文句子 + 段落末尾有紫色字体衍生词
  - **行 50**：
    - A 列（编号）：`48`
    - B 列（词源）：`Bacchus`
    - C 列（典故）：`Roman god of drama, wine, and ecstasy`
    - D 列（重点词汇）：`Bacchanalian a. 饮酒作乐的`
- **注意**：
  - 右侧的中文背景介绍（"酒神是古希腊及古罗马..."）**忽略**，不是结构化数据
  - 4 列严格分开（A=编号, B=词源, C=典故, D=重点词汇）
- **结果**：✅ 成功（首次使用神话典故模式）
- **2026-06-26 修正**：表头从"A 复用"改为"4 列分开"，Bacchus 行已重写

**测试 5（2026-06-26）— 4 张多组图批量写入（9 条：8 root + 1 prefix）**：
- **图 1**（root sheet，`ced, cess/grad, gress/vad, vas` 混合分割 → 3 条）：
  - 行 177 / 编号 175：`ced, cess | to go | antecedent, preceding, unprecedented, predecessor, concede | （留空）`
  - 行 178 / 编号 176：`grad, gress | to go | gradual, degrade, gradient, digressive | digressive=tangential; digressive speech`
  - 行 179 / 编号 177：`vad, vas | to go | invade, pervasive | television's pervasive influence; pervade; - Corruption pervades all levels of government in Greece.`
- **图 2**（root sheet，`curs, curr/ambul, ambl` 混合分割 → 2 条）：
  - 行 180 / 编号 178：`curs, curr | to run | precursor, current, currency, concurrent, cursive, discursive | give the letter a cursory reading; have a wide currency in China; discursive writing`
  - 行 181 / 编号 179：`ambul, ambl | to run | ambulance, ambulatory, amble | ambulatory population`
- **图 3**（prefix sheet，`per` 单条）：
  - 行 94 / 编号 92：`per | through | pervade, perceive, permeate, percolate, perspire, perennial, perspicacious | Per aspera, ad astra`
- **图 4**（root sheet，`spic, spec, spect/vis/scop` 复杂混合分割 → 3 条）：
  - 行 182 / 编号 180：`spic, spec, spect | to look, see | spectator, specter, perspicacious, circumspect, inspection, perspective, prospect | the specter of the ancient castle; the specter of earthquake; market prospect`
  - 行 183 / 编号 181：`vis | to look, see | vista, envisage | The economic vista for the next two years is excellent.; visionary a/n.`
  - 行 184 / 编号 182：`scop | to look, see | telescope, microscope, endoscope | （留空）`
- **总计**：root +8，prefix +1，共 9 条
- **踩坑**：第一次写入失败 `PermissionError`，原因是 Excel 仍占用文件。**必须先关闭 Excel 再写入**。

**测试 6（2026-06-26）— 列结构修正**：
- **问题**：mythology&history sheet 原表头为 A=词源（实际只放编号），B 列无表头，C=典故，D=重点词汇。A 列是合并的"编号+词源"格式（如 `48 Bacchus`）
- **用户要求**：4 列严格分开：A=编号, B=词源, C=典故, D=重点词汇
- **修复步骤**：
  1. 解除 A1:D1 标题行合并 → 重新合并 + 补"神话与历史词源"
  2. 解除 A2:B2 表头合并 → A2=`编号`, B2=`词源`
  3. C2/D2 保持 `典故`/`重点词汇`
  4. Bacchus 第 50 行重写：A=48, B=Bacchus, C=典故, D=重点词汇
- **结果**：✅ 成功
- **后续规则**：mythology&history sheet 写入前必须**自动验证表头**，确保是 4 列结构
