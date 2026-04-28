# 武汉东湖学院毕业论文自动化脚本原理与复用规范

本文件说明论文生成过程中可复用脚本的工作原理。它不是某一篇论文的专用脚本清单，而是给后续 agent 提供可迁移的自动化思路：哪些问题必须交给 Word 引擎处理，哪些格式必须用 OpenXML 固化，哪些步骤只能在渲染后验收。

可复用脚本本体位于本指南同级目录的 `scripts` 文件夹。后续 agent 应优先使用 `scripts` 中的参数化脚本，不要直接复用临时工作目录里带硬编码路径的脚本。

## 1. 为什么需要脚本

武汉东湖学院毕业论文格式中，封面、目录、页码和图表并不是普通文本排版问题。仅靠文字指南容易出现以下失败：

1. 封面缺少校徽、校名字样，或日期、英文封面被正文样式挤到下一页。
2. 目录使用手写 `……` 和静态页码，正文变动后页码失效。
3. Word 自动更新目录后重新写入默认缩进，导致目录视觉偏离范本。
4. 页码域、TOC 域在 `officecli`、LibreOffice 或 WPS 中显示不完整，需要 Microsoft Word 分页引擎更新。
5. 中文路径在 Windows PowerShell 5 中可能被非 UTF-8 解释，导致脚本打开错误路径。

因此，自动化脚本的核心任务不是“替代格式指南”，而是把指南中的规则变成可重复执行、可验收的 Word 操作。

## 2. 已验证脚本类型与原理

### 2.1 主生成脚本

代表脚本：`scripts/examples/build_thesis_draft_project_example.py`

原理：

1. 使用 `python-docx` 创建 `.docx`，按章节、段落、表格、图片等 Word 对象写入，而不是把论文拼成纯文本。
2. 显式设置页面尺寸、页边距、页眉页脚、正文样式、标题样式、图题表题样式。
3. 为标题样式写入 outline level，使 Word 能识别 `1`、`1.1`、`1.1.1` 等层级并生成目录。
4. 插入 Word 域代码，例如 `PAGE` 页码域和 `TOC` 目录域，而不是手写页码。
5. 图、表、代码截图等对象在生成阶段先放到正确章节，再由后处理负责更新目录和页码。

适用场景：

- 从项目资料生成初稿。
- 大范围重排论文。
- 需要统一正文、标题、图题、表题、参考文献样式。

注意：

- `python-docx` 能写入域代码和样式，但不能像 Word 一样完成真实分页和目录页码计算。
- 生成脚本结束后必须再进入 Word COM 更新字段。
- `examples` 下的生成脚本包含当前项目题目、学生信息、素材目录和正文内容，只能作为结构参考，不能直接套用到其他论文。

### 2.2 封面复制脚本

代表脚本：`scripts/copy_cover_from_sample.ps1`

原理：

1. 用 Microsoft Word COM 打开 `范本.docx` 和目标论文。
2. 通过 `GoTo(Page 3)` 找到范本第 3 页起点。
3. 取范本文档从开头到第 3 页起点之前的 `FormattedText`，即完整复制中文封面和英文封面。
4. 删除目标论文原有前两页，将范本前两页插入目标论文文首。
5. 只替换学号、题目、专业、学生、指导教师、日期等字段文本，不重建版面。
6. 再把范本前 47 个段落的直接段落格式重新应用到目标文档，防止目标文档 `Normal` 样式污染封面。

为什么这样做：

- 封面包含校徽、校名字样、图片锚点、空段、固定行距和分页节奏。
- 从零重建封面很容易缺图片或错页。
- 复制 `FormattedText` 能最大程度保留模板中不可见的 Word 排版信息。

适用场景：

- 项目中存在可信 `范本.docx`。
- 需要快速修复前两页与范本不一致的问题。

禁止事项：

- 不要把范本整页截图贴入论文。
- 不要删除模板空段后再手工排版。
- 不要覆盖范本原文件，必须在目标副本上处理。

### 2.3 目录更新与固化脚本

代表脚本：`scripts/update_toc_and_freeze.ps1`、`scripts/postprocess_toc_xml.py`、`scripts/verify_toc_xml.py`

该流程分为两步。

第一步：Word COM 更新动态字段。

1. 打开目标 `.docx`。
2. 执行 `doc.Repaginate()`，让 Word 重新分页。
3. 执行 `doc.Fields.Update()` 和 `doc.TablesOfContents.Item(i).Update()`。
4. 保存文档。

这一步必须由 Word 完成，因为目录页码取决于 Word 的分页引擎。

第二步：OpenXML 固化目录格式。

1. 解压 `.docx`，读取 `word/document.xml` 和 `word/styles.xml`。
2. 找到目录段落样式 `TOC1`、`TOC2`、`TOC3`。
3. 清除 Word 自动写入的目录缩进。
4. 写入固定行距、段前段后、右对齐制表位和点线前导符。
5. 重新打包为 `.docx`。

这一步必须在 Word 更新目录之后做。原因是 Word 更新目录时可能重新生成目录段落，并把默认缩进写回 XML。如果先改 XML 再更新目录，格式可能被 Word 覆盖。

## 3. 目录 OpenXML 推荐参数

目录项应具备以下 XML 结构或等效格式：

| 对象 | 推荐值 | 说明 |
|---|---|---|
| 段落样式 | `TOC1`、`TOC2`、`TOC3` | 不能是 `Normal` |
| 右对齐制表位 | `w:tab w:val="right"` | 页码靠右 |
| 点线前导符 | `w:leader="dot"` | 生成点线，不手写省略号 |
| 制表位位置 | `w:pos="8504"` 左右 | 约等于范本 425.2 pt |
| 行距 | `w:line="460"`，`w:lineRule="exact"` | 固定 23 磅 |
| 段前段后 | `w:before="0"`、`w:after="0"` | 避免目录稀疏 |
| 缩进 | `w:left="0"`、`w:firstLine="0"` | 与范本左边界一致 |

如果页面边距变化，`w:pos` 应按版心右边界重新换算，不得固定套用到所有论文。

## 4. 可复用执行流程

生成或重修论文时，推荐按以下顺序执行：

1. 用主生成脚本生成结构化论文：标题必须有 outline level，目录必须是 TOC 域。
2. 用 `verify_front_matter_layout.py` 检查独创性声明、授权书、中文摘要、英文摘要和目录的页序；此时允许失败，但必须先看清失败项。
3. 如果有 `范本.docx`，用封面复制流程替换前两页。
4. 再次用 `verify_front_matter_layout.py` 检查前置页，确认封面复制没有破坏第 3 页以后的分页。
5. 用 Word COM 打开目标文档，更新所有字段、目录和页码。
6. 用 OpenXML 后处理固化目录段落格式。
7. 用 `officecli validate <docx>` 验证文档结构。
8. 用 `officecli view <docx> outline` 检查标题层级。
9. 解包检查目录段落：不得有硬编码 `……`，`TOC1/TOC2/TOC3` 的缩进、行距、制表位必须符合规范。
10. 导出 PDF/PNG，与范本页面做视觉核对。

## 4.1 scripts 目录调用示例

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\copy_cover_from_sample.ps1" `
  -SamplePath ".\范本.docx" `
  -DraftPath ".\论文初稿.docx" `
  -OutputPath ".\论文初稿_封面修正版.docx" `
  -ReplacementsJson ".\cover_replacements.json"

powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\update_toc_and_freeze.ps1" `
  -DocxPath ".\论文初稿_封面修正版.docx" `
  -PythonExe "python"

python ".\scripts\verify_toc_xml.py" ".\论文初稿_封面修正版.docx"

python ".\scripts\verify_cover_identifiers_and_length.py" ".\论文初稿_封面修正版.docx" `
  --student-id "2022040731173"

python ".\scripts\verify_front_matter_layout.py" ".\论文初稿_封面修正版.docx" `
  --title-cn "中文论文题目" `
  --title-en "English Thesis Title"
```

当修改指南、脚本或生成器后，应执行全量格式回归测试：

```powershell
python ".\scripts\run_full_format_regression.py"
```

该测试会从示例生成脚本开始，依次执行封面复制、目录更新、OpenXML 后处理、前置页检查、右上角学号/档号与字数检查、目录检查、图表检查、`officecli validate` 和 `officecli view outline`。如果 `officecli validate` 输出 schema 错误，即使命令退出码为 0，也必须视为失败继续修复。

`cover_replacements.json` 使用数组格式：

```json
[
  {"find": "模板中的原文字", "replace": "替换后的文字"}
]
```

## 5. 脚本复用安全规则

1. 复用脚本时必须把输入路径、输出路径、范本路径改为参数，不要保留某次任务的硬编码路径。
2. Windows PowerShell 5 可能按本地编码读取脚本；涉及中文路径、`宋体`、`黑体`、`目　录` 等中文字符串时，优先使用参数传入或 UTF-8 Base64 解码。
3. 不得直接修改 `范本.docx`，必须复制到目标文件或临时文件后处理。
4. 不得在 Word 已打开同一文件时运行脚本；Windows 文件锁会导致保存失败或验证失败。
5. `officecli`、Word COM、Python OpenXML 脚本不要并行操作同一个 `.docx`。
6. 每次脚本后处理后必须重新验证，不能根据脚本退出码直接声明完成。

## 6. 最小验收命令

推荐每次自动化处理后执行：

```powershell
officecli validate "<论文.docx>"
officecli view "<论文.docx>" outline
```

目录页还应执行 OpenXML 检查，至少确认：

```text
TOC 段落数 > 0
目录样式包含 TOC1/TOC2/TOC3
BAD_IND = 0
BAD_SPACING = 0
BAD_TABS = 0
```

封面页还应导出 PDF/PNG，确认：

```text
第 1 页有校徽和校名字样
中文日期仍在第 1 页
英文封面从第 2 页开始
独创性声明从第 3 页开始
前两页无正文页眉和正文页码
```

右上角学号/档号与字数还应执行 OpenXML 检查，至少确认：

```text
COVER_IDENTIFIERS_AND_LENGTH_CHECK=PASS
NONSPACE_VISIBLE_CHARS >= 20000
学号段保留 8280 twips 制表位和冒号后的单下划线
档号段保留 7920、8280 twips 制表位；无档号时仍有下划线占位
```

声明、授权书和摘要页还应执行 Word COM 页序检查，至少确认：

```text
FRONT_MATTER_CHECK=PASS
所有节的 docGrid type = linesAndChars
所有节的 docGrid linePitch = 312
独创性声明在第 3 页
学位论文版权使用授权书在第 3 页
第 3 页不包含中文题目、摘要或关键词
中文题目、摘　要、关键词在第 4 页
英文题目、ABSTRACT 在中文摘要后单独起页
Key words 在目录之前
目录在英文摘要之后
```

如果视觉上仍然比范本“挤”，不要只看 `LineSpacingRule`。还必须检查：

```text
空段的段落标记字号：声明正文后的空段应为 14 pt
w:pPr/w:rPr/w:sz 与 w:szCs 是否同时写入
docGrid linePitch 是否被 python-docx 默认写成 360
```

## 7. 给后续 agent 的判断规则

1. 若问题是“内容怎么写”，先看 `论文生成总流程清单.md`。
2. 若问题是“页面应该长什么样”，先看主格式指南、封面规范、目录规范。
3. 若问题是“为什么按指南写了仍然不对”，优先查看本文件，因为问题很可能出在 Word 自动字段、目录重生成、模板直接格式或文件锁。
4. 若要复用当前项目脚本，必须先抽掉硬编码路径，再在副本上试运行。
5. 若脚本输出与范本视觉不一致，以渲染后的 PDF/PNG 为准继续修正。
