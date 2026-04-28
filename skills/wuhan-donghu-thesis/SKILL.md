---
name: wuhan-donghu-thesis
description: Use when generating, editing, formatting, or validating Wuhan Donghu College undergraduate thesis Word drafts, especially .docx papers needing school-specific cover pages, declarations, abstracts, TOC, figures, tables, references, page numbers, or template comparison.
---

# Wuhan Donghu Thesis

## Purpose

Use this skill for Wuhan Donghu College undergraduate thesis `.docx` drafts. The school format is fragile: cover pages, front matter, TOC fields, page numbering, figures, and three-line tables must be handled as Word objects, not plain text.

If the host platform provides DOCX and process helpers, use them. In Codex, prefer `officecli-docx` for `.docx` inspection/editing and Superpowers process skills when they apply. On other platforms, follow this skill's references and bundled scripts directly.

## First Steps

1. Read the project’s `AGENTS.md` if present.
2. Read `references/word-format.md` for global Word rules.
3. Read `references/generation-workflow.md` before writing thesis content from task books, proposals, source code, screenshots, or project materials.
4. If touching the first two pages, read `references/cover-pages.md`.
5. If touching the table of contents, read `references/toc-page.md`.
6. If using scripts or Word COM/OpenXML, read `references/automation.md`.
7. If installing or running outside Codex, read `references/platform-compatibility.md`.

## Platform Compatibility

This skill uses the Agent Skills folder pattern: `SKILL.md` is the entry point, while `references/`, `scripts/`, and `assets/` are relative bundled resources. `agents/openai.yaml` is optional OpenAI/Codex UI metadata and can be ignored by other platforms.

Full automation requires Windows with Microsoft Word COM. Non-Windows agents can still use the references and OpenXML validators, but final TOC/page-number refresh and template-faithful cover copying should be run in Word or by a Windows worker.

## Core Workflow

1. Build an evidence ledger from task documents, proposal, source code, screenshots, database/schema, and figure/table references.
2. Generate a structured `.docx`: real Word sections, paragraph styles, heading outline levels, PAGE fields, and TOC field. Do not handwrite page numbers or TOC dot leaders.
3. Treat cover pages as template-level content. Prefer copying `范本.docx` first two pages with `scripts/copy_cover_from_sample.ps1`; only replace fields. The Chinese cover's top-right 学号/档号 lines and the 院系/专业/姓名/教师 lines must preserve fixed-width underlined slots; short values must be centered/padded with underlined spaces so the right edge stays aligned. 档号 may be blank only as a full-width underlined placeholder, never as a deleted line.
4. Generate enough content for the school length rule: final visible non-space text must be at least 20000 characters; target 22000-25000 for drafts to leave editing margin.
5. Update TOC and page numbers with Microsoft Word, then freeze TOC/table XML with `scripts/update_toc_and_freeze.ps1`.
6. Validate front matter, cover identifiers, length, TOC, figures, tables, OpenXML schema, and outline before saying the draft is ready.

## Bundled Scripts

Resolve script paths relative to this skill directory.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\copy_cover_from_sample.ps1" `
  -SamplePath ".\范本.docx" `
  -DraftPath ".\论文初稿.docx" `
  -OutputPath ".\论文初稿_封面修正版.docx" `
  -ReplacementsJson ".\cover_replacements.json"

powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\update_toc_and_freeze.ps1" `
  -DocxPath ".\论文初稿_封面修正版.docx" `
  -PythonExe "python"

python ".\scripts\verify_front_matter_layout.py" ".\论文初稿_封面修正版.docx" `
  --title-cn "中文论文题目" `
  --title-en "English Thesis Title"

python ".\scripts\verify_toc_xml.py" ".\论文初稿_封面修正版.docx"
python ".\scripts\verify_cover_identifiers_and_length.py" ".\论文初稿_封面修正版.docx" `
  --student-id "2022040731173"
python ".\scripts\verify_cover_against_sample.py" `
  --sample ".\范本.docx" `
  --candidate ".\论文初稿_封面修正版.docx" `
  --student-id "2022040731173"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\verify_cover_fonts_against_sample.ps1" `
  -SamplePath ".\范本.docx" `
  -CandidatePath ".\论文初稿_封面修正版.docx"
python ".\scripts\verify_figure_table_layout.py" ".\论文初稿_封面修正版.docx"
officecli validate ".\论文初稿_封面修正版.docx"
officecli view ".\论文初稿_封面修正版.docx" outline
```

For project-specific diagram checks:

```powershell
python ".\scripts\verify_figure_table_layout.py" ".\论文初稿_封面修正版.docx" `
  --build-script ".\生成论文脚本.py" `
  --require-er-function `
  --forbidden-er-token "users" `
  --check-api-flow-blank-space
```

## Validation Rules

Fresh verification is mandatory. Minimum evidence:

- `verify_front_matter_layout.py`: `FRONT_MATTER_CHECK=PASS`
- `verify_cover_identifiers_and_length.py`: `COVER_IDENTIFIERS_AND_LENGTH_CHECK=PASS`, `NONSPACE_VISIBLE_CHARS >= 20000`
- `verify_cover_against_sample.py`: `COVER_SAMPLE_COMPARISON=PASS` when `范本.docx` is available; this includes fixed-width underlined slot checks for 学号/档号 and 院系/专业/姓名/教师
- `verify_cover_fonts_against_sample.ps1`: `COVER_FONT_COMPARISON=PASS` when Windows Word is available; this catches theme/font pollution such as Cambria, MS Mincho, `+西文正文`, or `+中文正文` replacing the sample's fonts
- `verify_toc_xml.py`: `BAD_IND=0`, `BAD_SPACING=0`, `BAD_TABS=0`, `HARDCODED_LEADERS=0`
- `verify_figure_table_layout.py`: `PASS: figure/table layout checks passed`
- `officecli validate`: no schema errors in output
- `officecli view ... outline`: heading tree includes expected 1 / 1.1 / 1.1.1 levels

If `officecli validate` prints schema errors while returning exit code 0, treat it as failure.

## Common Failure Modes

- Cover pages rebuilt as normal paragraphs: use template copy or bundled cover assets.
- TOC manually typed with `……`: replace with a Word TOC field and right-tab dot leaders.
- Word updates TOC then reintroduces default indentation: run OpenXML postprocessing after Word update.
- Tables look correct but XML is invalid: preserve schema order when inserting `tblW`, `tblBorders`, `tcW`, `tcBorders`, paragraph `tabs/spacing/ind/jc`, and run `rFonts/sz`.
- Table validation only checks database fields: validate every table in the whole paper.
- Current-project facts leak into a different thesis: change title, student info, cover replacements, source evidence, schema, screenshots, and figure labels.

## Assets

Use `assets/wdu_cover_logo.png` and `assets/wdu_cover_calligraphy.png` only when no `范本.docx` is available. Reference images are visual checks, not full-page screenshots to paste into the paper.
