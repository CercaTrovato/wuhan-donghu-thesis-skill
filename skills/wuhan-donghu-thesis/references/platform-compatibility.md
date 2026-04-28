# Platform Compatibility

Use this reference when installing or running the skill outside the current Codex desktop environment.

## Portable Skill Contract

The skill is portable as a folder, not as a single prompt. Keep these paths together:

- `SKILL.md`: required entry point.
- `references/`: detailed thesis rules and workflows.
- `scripts/`: deterministic Word/OpenXML helpers and validators.
- `assets/`: Wuhan Donghu cover logo/calligraphy and visual cover references.
- `agents/openai.yaml`: optional OpenAI/Codex metadata; safe to ignore elsewhere.

Run every script path relative to the skill directory unless the host platform copies resources elsewhere.

## GitHub Distribution

For GitHub publishing, package this skill as a repository with the real skill at:

```text
skills/wuhan-donghu-thesis/
```

Keep platform metadata and installers at the repository root:

- `README.md`: platform install guide and one-line install commands.
- `install.ps1`: Windows PowerShell installer.
- `install.sh`: macOS/Linux/WSL installer.
- `.codex-plugin/plugin.json`: Codex plugin metadata.
- `.opencode/plugins/*.js`: OpenCode plugin entry that registers `skills/`.
- `.opencode/INSTALL.md`: OpenCode-specific install instructions.
- `.claude-plugin/plugin.json`, `.cursor-plugin/plugin.json`, `gemini-extension.json`, `GEMINI.md`: optional platform metadata.

One-line installers should download the GitHub archive, extract `skills/wuhan-donghu-thesis`, and copy it into the target platform's skill directory. Do not ask users to copy individual reference, script, or asset files by hand.

## Codex

Install by placing the folder at:

```powershell
%USERPROFILE%\.codex\skills\wuhan-donghu-thesis
```

Codex should discover the skill from `SKILL.md`. When available, use `officecli-docx` for `.docx` inspection and Superpowers process skills for planning, debugging, verification, and skill updates.

## OpenCode

Install by placing the same folder in the OpenCode skills directory used by the local setup. In this workspace, that directory is:

```powershell
E:\AIworkspace\opencode_workspace\skills\wuhan-donghu-thesis
```

If OpenCode does not auto-discover Agent Skills folders in a given environment, manually load `SKILL.md` first, then load only the relevant files under `references/` for the current operation.

## Generic AI Automation Platforms

For platforms without native skill discovery:

1. Upload or mount the whole `wuhan-donghu-thesis` folder.
2. Instruct the agent to read `SKILL.md` before any thesis work.
3. Instruct the agent to resolve linked references and scripts relative to that folder.
4. Provide a shell or external worker for PowerShell, Python, and Microsoft Word tasks.

The minimum usable context is `SKILL.md` plus the relevant reference file. The full reliable workflow also needs the scripts and assets.

## Runtime Requirements

Required for full automation:

- Windows.
- Microsoft Word desktop app with COM automation enabled.
- PowerShell 5+ or PowerShell 7+.
- Python 3.10+.
- Python packages: `lxml`; `pywin32` for Word COM validators.

Recommended:

- `officecli` for `.docx` schema validation and outline inspection.
- A sample `范本.docx` when copying the first two cover pages exactly.

Not required for normal skill use:

- Pandoc. It is only useful when converting Markdown guide files to `.docx` documentation.

## Non-Windows Fallback

Non-Windows platforms can still:

- Generate thesis text and section structure.
- Inspect and patch `.docx` OpenXML.
- Run validators that do not require Word COM, such as `verify_toc_xml.py` and most of `verify_figure_table_layout.py`.

Non-Windows platforms cannot reliably:

- Refresh Word TOC page numbers exactly as Word displays them.
- Copy the first two cover pages with the same fidelity as Word COM.
- Verify rendered front matter pagination through Word.

For final delivery, hand the `.docx` to a Windows + Word worker and run the full validation checklist from `SKILL.md`.

## Portability Limits

This skill is suitable for common agent platforms that can load local files and run shell commands. It is not a one-click universal plugin for every AI platform because each platform has different skill discovery, file mounting, and tool execution rules. Treat `SKILL.md` as the canonical workflow and the bundled files as the implementation package.
