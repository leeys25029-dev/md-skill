---
name: weekly-report
description: Convert Markdown paper reading notes (Obsidian notes) into strictly formatted docx weekly reports. Use this skill when the user wants to turn literature reading notes (.md) into a Word weekly report (周报) matching a fixed template — triggers include "整理成周报", "生成周报", "把笔记转成周报 docx", "weekly report from notes".
description_zh: 把 Markdown 文献阅读笔记转化为严格符合模板格式的 docx 周报。当用户要把论文阅读笔记整理成周报、生成周报 Word、把 md 笔记按周报模板排版时使用。
description_en: Convert Markdown paper reading notes into strictly formatted docx weekly reports matching a fixed template. Use when turning literature notes into a Word weekly report.
version: "1.0.0"
agent_created: true
---

# 周报生成（Markdown 笔记 → docx 周报）

把 Obsidian 等来源的 Markdown 文献阅读笔记转化为严格符合固定模板格式的 docx 周报。格式通过内置模板 + 生成脚本保证——**禁止手工用 python-docx 从零造样式，必须走模板脚本流程**。

## 触发场景

- 用户给一个或多个 md 笔记文件，要求整理成周报 / 生成周报 Word / 按周报模板排版
- 用户提到"周报""weekly report"且素材是文献阅读笔记

## 资源

本 skill 目录（`<skill_dir>`）下：

| 文件 | 用途 |
|------|------|
| `scripts/build_report.py` | 周报生成脚本（唯一格式出口） |
| `assets/weekly-report-template.docx` | 周报模板（含全部命名样式，勿改动） |
| `assets/sample-weekly-report.docx` | 原版示范周报（对照参考） |
| `references/format-spec.md` | 周报格式规范明细（存疑时查阅） |

`<skill_dir>` 即本 SKILL.md 所在目录。脚本内默认模板路径已指向 `../assets/weekly-report-template.docx`，无需额外传 `--template`。

## 工作流程

### 第一步：收集输入

需要的信息（用户消息里已有就直接用，缺失则一次性问全，不要拆成多轮）：

1. **笔记文件路径**（必需）：一个或多个 md 文件
2. **周报日期**（必需）：大标题格式为 `YYYY-MM-DD 周报`；用户未给时询问
3. **输出位置**：默认 `~/Downloads/{日期}周报.docx`，用户可指定
4. **实验汇报小节标题**（可选）：默认「实验汇报」；用户有具体实验主题（如「在DLinear模型上添加树突模块」）就用用户的标题
5. **文献调研起始编号**（可选）：用户若维护连续论文编号（如从 [63] 起）则用其起始号；未指定则从 [1] 开始

### 第二步：解析每篇笔记

读取全部 md 笔记，按下表映射提取内容（笔记结构与映射不完全匹配时灵活处理，缺哪部分就跳过哪部分，不要编造）：

| 笔记内容 | 周报去向 |
|---------|---------|
| 基本信息表（论文标题/年份/会议） | 提取论文短标题用于二级标题；标题/会议/年份用于文献调研条目；表格本身不进周报 |
| 「论文目的 / 结构扫描」 | `2.X` 下的 **文章动机**（三级标题）正文 |
| 「具体的方法」下各 `###` 小节 | **文章架构**（三级标题）之后，每个小节一个三级标题 + 正文/编号列表/图片 |
| 「实验结果」小节 | **实验结果**（三级标题） |
| 「一句话总结 / 总体评价 / 对我的研究有什么帮助」 | 素材，用于提炼 `1.X` 的主要内容摘要 |
| 「下一步行动」清单 | `3. 下一步安排` 的素材（与用户口述的安排合并，用户口述优先） |

图片与语法清理规则：
- `![[Screenshot xxx.png|187]]` → 图片块，宽度 = 187÷96 ≈ 1.95 英寸；`|N` 缺省时宽度取 4.5 英寸；上限 6.8 英寸
- 图片插入对应原文位置：紧跟其所属正文之后（示范周报中图片常与说明文字同段，说明文字用 `web` 类型段落）
- `[[双链笔记名]]` → 去掉双方括号，仅保留链接文字
- `#标签` → 删除；`- [ ]` / `- [x]` 复选框 → 去掉符号保留文字；`**加粗**` → 去掉星号保留文字

### 第三步：撰写内容，生成 spec.json

周报整体结构（固定骨架，逐块写入 `blocks` 数组）：

```
大标题: {日期} 周报                    （脚本自动加，spec 提供 date 即可）
1.主要内容
  1.1 ~ 1.N  每篇文献标题（h2）+ 一段摘要（para）
  1.N+1      实验汇报（h2）+ 灰色占位（placeholder）
  [文献调研]                          ← 三级标题，文本固定为 "[文献调研]"
             [编号] 引用条目若干（para_noindent，每条一段，自动生成）
2.具体内容
  2.1 ~ 2.N  每篇文献：文章动机（h3）→正文；文章架构（h3）→
             各方法小节（h3）→正文/编号列表/图片；实验结果（h3）
  2.N+1      实验汇报（h2）+ 灰色占位（placeholder）
3.下一步安排
  编号列表若干（para_noindent，"1. xxx" "2. xxx"）
```

各部分写作规范：

- **1.X 摘要**（每篇 80~120 字，一个自然段）：套用「针对…问题，X 提出/采用…方法，…（机制一句话），（效果/评价一句话）」的句式，从笔记的论文目的、方法、总结提炼，客观陈述，不堆砌形容词
- **2.X 详细内容**：**保留笔记原文的口语化表达**，只做轻度清理（修正明显笔误、去掉笔记专属语法、合并断句），不要重写、不要书面化拔高——示范周报的正文就是笔记原文的直接搬运
- **编号列表项**（"1. 序列经过快速傅里叶变换…"）：用 `para_noindent` 类型（首行不缩进），编号数字写在文本里
- **图片说明文字**（如"复数乘法：长度相乘，角度相加"）：用 `web` 类型段落，可与图片放同段（runs 混排）
- **实验汇报占位**：`placeholder` 类型，文本固定为「（此处由本人填写本周实验进展，完成后删除此行）」
- **文献调研**：在「1.主要内容」章节末尾（实验汇报之后、「2.具体内容」之前）生成三级标题，文本固定为 `[文献调研]`，其后**自动生成引用条目**——每篇笔记一条，用 `para_noindent` 类型，一段一条，格式严格参照示范周报：
  ```
  [编号] 姓 名首字母, 姓 名首字母, 姓 名首字母, et al. 论文标题[文献类型]//出处 年份.
  ```
  - **编号**：按第一步收集的起始编号连续递增
  - **作者**：笔记基本信息表无作者字段，**必须联网搜索补全**（按论文标题 + 会议年份搜 NeurIPS/dblp/OpenReview/arXiv/Google Scholar），取前 3 位作者 + et al.；中文姓名按「姓 + 名首字母」缩写（如 Xu-Yang Chen → Chen X）；搜不到则在作者位置写「（待补作者）」
  - **文献类型**：会议论文用 `[C]//`，期刊论文用 `[J].`
  - **出处**：从笔记基本信息表的「期刊/会议」字段取
  - **标题**：用笔记基本信息表的论文标题（英文原标题，保持原大小写）
  - 示例：`[1] Lu H, Chen X, Ye H, et al. SOFTS: Efficient multivariate time series forecasting with series-core fusion[C]//NeurIPS 2024.`
- **h1/h2 的 text 不要带编号**，脚本自动编号（h1 → "N.标题"，h2 → "N.M 标题"）

spec.json 写入临时文件（如 `/tmp/weekly_spec.json`），块类型速查：
`h1` / `h2` / `h3` / `para`（正文缩进）/ `para_noindent`（列表）/ `web`（Normal Web）/ `image`（独立图片段）/ `placeholder`（灰色占位）；`text` 与 `runs` 二选一，`runs` 支持文字与 `{"image": "文件名", "width": 英寸}` 混排。

完整 spec.json 结构见 `scripts/build_report.py` 顶部注释。

### 第四步：运行脚本生成 docx

```bash
python3 <skill_dir>/scripts/build_report.py /tmp/weekly_spec.json \
  --output <输出路径>.docx \
  --vault "<用户的Obsidian库根目录>"
```

- `--vault` 默认为 `~/Library/Mobile Documents/iCloud~md~obsidian/Documents/Tang's vault`，笔记不在该库时按笔记路径推断或询问用户
- 脚本会自动编号标题、应用命名样式、解析并插入图片；缺失 python-docx 时会自动安装

### 第五步：验证与交付

1. 检查脚本输出：如有 `⚠️ 图片未找到` 警告，向用户列出未找到的图片并说明已在文中以红色占位标注
2. 用 python-docx 快速抽查输出文件：大标题/一级/二级标题样式、二级标题颜色 #0F4761、段落数量与预期一致
3. 用 present_files 向用户交付 docx，并简报：本篇周报包含几篇文献、实验部分已预留、文献调研条目数、未找到的图片（如有）

## 格式规范

完整格式明细见 `references/format-spec.md`。关键点：

- 大标题：Times New Roman 22pt 居中加粗
- 一级标题：仿宋+Times New Roman Bold 15pt 黑色
- 二级标题：仿宋+Times New Roman Bold 14pt **深蓝 #0F4761**
- 三级标题：仿宋+Times New Roman 12pt 加粗
- 正文：仿宋+Times New Roman 12pt，1.5 倍行距，首行缩进 2 字符；编号列表首行不缩进
- A4，上下边距 72pt，左右 90pt

## 注意事项

- **禁止**绕过脚本手工构造 docx 样式；**禁止**修改模板文件
- 用户后续要求「改内容重新生成」：修改 spec.json 重跑脚本即可；要求「改格式」（如换字号颜色）属于模板变更，先与用户确认再改模板并同步更新 references/format-spec.md
- 用户只给了一篇笔记也要生成完整骨架（实验占位、文献调研、下一步安排照常）
- 下一步安排如果笔记和用户都没提供素材，生成灰色占位「（此处由本人填写下一步安排）」
- 涉及用户笔记内容时保持原意，不替用户下研究结论
- 所有回复使用简体中文
