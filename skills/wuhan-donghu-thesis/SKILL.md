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

1. Build an evidence ledger from task documents, proposal, source code, screenshots, database/schema, and figure/table references. The ledger must include implemented API routes, frontend views, model/table names, AI inputs/outputs, and any missing or unverifiable functions.
2. Generate a structured `.docx`: real Word sections, paragraph styles, heading outline levels, PAGE fields, and TOC field. Do not handwrite page numbers or TOC dot leaders.
3. Treat cover pages as template-level content. Prefer copying `范本.docx` first two pages with `scripts/copy_cover_from_sample.ps1`; only replace fields. The Chinese cover's top-right 学号/档号 lines and the 院系/专业/姓名/教师 lines must preserve fixed-width underlined slots using Word-rendered x coordinates, not character counts; short values must be centered/padded with underlined spaces so the visual left and right edges match the sample. 档号 may be blank only as a full-width underlined placeholder, never as a deleted line.
4. Generate enough body content for the school length rule: the thesis body from `1 绪论` through the last body chapter before `参考文献` must be at least 20000 visible non-space characters. Do not count cover pages, declarations, abstracts, TOC, references, or acknowledgments toward this requirement. Target 22000-25000 body characters for drafts to leave editing margin.
5. Draw structure diagrams, flowcharts, module diagrams, deployment diagrams, and E-R diagrams in the reference style: pure white background, black text, black lines, no gray fill, no accent color, no grid, no dark theme. Use one font family consistently across generated diagrams.
6. The total E-R diagram must cover all implemented system themes found in the evidence ledger. Distinguish physical tables from logical/business themes in the nearby text instead of omitting logical themes. For total E-R diagrams, emit an `er_layout.json` manifest and reject layouts with tangled lines, shared ports, connector crossings, relation diamonds overlapping boxes, or `1/n` labels inside boxes.
7. Code figures should look like clean IDE/editor screenshots when generated from snippets: white editor background, line numbers or editor chrome, readable monospaced code, and CJK font fallback for Chinese strings. Emit a `code_figures_manifest.json` and reject any source text or configured font set that can render Chinese as mojibake, boxes, or question marks.
8. Update TOC and page numbers with Microsoft Word, then freeze TOC/table XML with `scripts/update_toc_and_freeze.ps1`.
9. Validate front matter, cover identifiers, body length, TOC, figures, tables, diagram colors, code figure text, project evidence coverage, OpenXML schema, and outline before saying the draft is ready.

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
python ".\scripts\verify_body_text_length.py" ".\论文初稿_封面修正版.docx" `
  --min 20000
python ".\scripts\verify_cover_against_sample.py" `
  --sample ".\范本.docx" `
  --candidate ".\论文初稿_封面修正版.docx" `
  --student-id "2022040731173"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\verify_cover_rendered_alignment.ps1" `
  -SamplePath ".\范本.docx" `
  -CandidatePath ".\论文初稿_封面修正版.docx"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\verify_cover_fonts_against_sample.ps1" `
  -SamplePath ".\范本.docx" `
  -CandidatePath ".\论文初稿_封面修正版.docx"
python ".\scripts\verify_figure_table_layout.py" ".\论文初稿_封面修正版.docx"
python ".\scripts\verify_monochrome_diagrams.py" ".\论文初稿_封面修正版.docx"
python ".\scripts\verify_project_content_completeness.py" ".\论文初稿_封面修正版.docx" `
  --requirements ".\project_requirements.json"
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

python ".\scripts\verify_monochrome_diagrams.py" ".\论文初稿_封面修正版.docx"
python ".\scripts\verify_project_content_completeness.py" ".\论文初稿_封面修正版.docx" `
  --requirements ".\scripts\examples\project_requirements_example.json"
python ".\scripts\verify_er_layout_geometry.py" ".\generated\er_layout.json"
python ".\scripts\verify_code_figure_text.py" ".\generated\code_figures_manifest.json"
```

`verify_project_content_completeness.py` is intentionally data-driven. For a new thesis, create a fresh requirements JSON from that project's source code instead of reusing another project's API list or ER themes.

## Validation Rules

Fresh verification is mandatory. Minimum evidence:

- `verify_front_matter_layout.py`: `FRONT_MATTER_CHECK=PASS`
- `verify_cover_identifiers_and_length.py`: `COVER_IDENTIFIERS_AND_LENGTH_CHECK=PASS` for cover structure and a whole-document sanity count
- `verify_body_text_length.py`: `BODY_TEXT_LENGTH_CHECK=PASS`, `BODY_NONSPACE_VISIBLE_CHARS >= 20000`; for robust drafts prefer `--min 22000`
- `verify_cover_against_sample.py`: `COVER_SAMPLE_COMPARISON=PASS` when `范本.docx` is available; this checks paragraph order, tab stops, label/suffix formatting, underlined runs, and preserved spaces
- `verify_cover_rendered_alignment.ps1`: `COVER_RENDERED_ALIGNMENT=PASS` when Windows Word is available; this checks the real rendered x coordinates of 学号/档号 and 院系/专业/姓名/教师 underline slots against `范本.docx`
- `verify_cover_fonts_against_sample.ps1`: `COVER_FONT_COMPARISON=PASS` when Windows Word is available; this catches theme/font pollution such as Cambria, MS Mincho, `+西文正文`, or `+中文正文` replacing the sample's fonts
- `verify_toc_xml.py`: `BAD_IND=0`, `BAD_SPACING=0`, `BAD_TABS=0`, `HARDCODED_LEADERS=0`
- `verify_figure_table_layout.py`: `PASS: figure/table layout checks passed`
- `verify_monochrome_diagrams.py`: `MONOCHROME_DIAGRAM_CHECK=PASS`; all structural/process/E-R diagrams must be white-background, black-only diagrams
- `verify_project_content_completeness.py`: `PROJECT_CONTENT_COMPLETENESS=PASS` with a project-specific requirements JSON generated from source evidence
- `verify_er_layout_geometry.py`: `ER_LAYOUT_GEOMETRY_CHECK=PASS` for generated total E-R diagram manifests
- `verify_code_figure_text.py`: `CODE_FIGURE_TEXT_CHECK=PASS` for generated code figure manifests containing non-ASCII strings
- `officecli validate`: no schema errors in output
- `officecli view ... outline`: heading tree includes expected 1 / 1.1 / 1.1.1 levels

If `officecli validate` prints schema errors while returning exit code 0, treat it as failure.

## Common Failure Modes

- Cover pages rebuilt as normal paragraphs: use template copy or bundled cover assets.
- TOC manually typed with `……`: replace with a Word TOC field and right-tab dot leaders.
- Word updates TOC then reintroduces default indentation: run OpenXML postprocessing after Word update.
- Tables look correct but XML is invalid: preserve schema order when inserting `tblW`, `tblBorders`, `tcW`, `tcBorders`, paragraph `tabs/spacing/ind/jc`, and run `rFonts/sz`.
- Table validation only checks database fields: validate every table in the whole paper.
- Diagrams look close in Word but use gray fills or mixed fonts: run `verify_monochrome_diagrams.py` and re-render PDF/PNG for visual confirmation.
- E-R diagrams look complete but visually tangled: generate a geometry manifest, keep each relationship on its own connection ports, leave visible line length between boxes and diamonds, and run `verify_er_layout_geometry.py` before inserting the image.
- Code figures show boxes or `????` in Chinese strings: render with a CJK-capable fallback font, generate `code_figures_manifest.json`, run `verify_code_figure_text.py`, then inspect the rendered PDF/PNG page.
- XML checks pass while a figure caption is split onto the next page: set picture paragraph `keep_with_next=True` and `keep_together=True`, set caption paragraph `keep_together=True`, then inspect rendered pages.
- Total E-R diagrams become too simple because only physical tables are drawn: include all source-backed business themes and explain which nodes are logical themes rather than tables.
- Generated content misses implemented functions: build and validate a project-specific requirements JSON listing API routes, views, models/tables, AI features, and required ER themes.
- Current-project facts leak into a different thesis: change title, student info, cover replacements, source evidence, schema, screenshots, and figure labels.

## Assets

Use `assets/wdu_cover_logo.png` and `assets/wdu_cover_calligraphy.png` only when no `范本.docx` is available. Reference images are visual checks, not full-page screenshots to paste into the paper.
