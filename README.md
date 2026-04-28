# Wuhan Donghu Thesis Skill

Wuhan Donghu College undergraduate thesis skill for generating, formatting, and validating `.docx` drafts against the school template. It includes the thesis workflow, format references, cover assets, and Word/OpenXML validation scripts.

The repository follows the same distribution shape used by cross-agent skill/plugin projects such as Superpowers: platform metadata lives at the repository root, while the real skill lives under `skills/wuhan-donghu-thesis/`.

## One-Line Install

Default repository slug used by the installer:

```text
CercaTrovato/wuhan-donghu-thesis-skill
```

After publishing under another GitHub account or repository name, either update `install.ps1` / `install.sh`, or set `WDU_THESIS_SKILL_REPO=owner/repo` in the command.

### Windows PowerShell

Codex default install:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.ps1 | iex"
```

OpenCode install:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$env:WDU_THESIS_SKILL_PLATFORM='opencode'; irm https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.ps1 | iex"
```

Install into all common local skill directories:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$env:WDU_THESIS_SKILL_PLATFORM='all'; irm https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.ps1 | iex"
```

### macOS / Linux / WSL

Codex default install:

```bash
curl -fsSL https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.sh | bash
```

OpenCode install:

```bash
curl -fsSL https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.sh | WDU_THESIS_SKILL_PLATFORM=opencode bash
```

Install into all common local skill directories:

```bash
curl -fsSL https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.sh | WDU_THESIS_SKILL_PLATFORM=all bash
```

## Platform Install Notes

### Codex

The installer copies the skill to both common Codex locations for compatibility:

- `~/.agents/skills/wuhan-donghu-thesis`
- `~/.codex/skills/wuhan-donghu-thesis`

The repository also includes `.codex-plugin/plugin.json`, so it can be packaged as a Codex plugin later.

### OpenCode

Preferred plugin install through `opencode.json`:

```json
{
  "plugin": ["wuhan-donghu-thesis@git+https://github.com/CercaTrovato/wuhan-donghu-thesis-skill.git"]
}
```

Restart OpenCode after editing the config. The plugin entry registers `skills/wuhan-donghu-thesis` automatically.

### Claude Code / Cursor / Generic Agents

The one-line installer supports `WDU_THESIS_SKILL_PLATFORM=claude`, `cursor`, or `all`. For platforms without native skill discovery, upload or mount this repository and instruct the agent to read `skills/wuhan-donghu-thesis/SKILL.md` first.

### Gemini CLI

```bash
gemini extensions install https://github.com/CercaTrovato/wuhan-donghu-thesis-skill
```

## Runtime Requirements

Full thesis automation still requires Windows + Microsoft Word desktop COM for high-fidelity cover copying, TOC/page-number refresh, and front-matter pagination validation. See `skills/wuhan-donghu-thesis/references/platform-compatibility.md` for the exact requirements and non-Windows fallback.

## Contents

- `skills/wuhan-donghu-thesis/`: skill entry point, references, scripts, and assets.
- `install.ps1` / `install.sh`: one-line installers for common local skill directories.
- `.opencode/plugins/wuhan-donghu-thesis.js`: OpenCode plugin that registers the bundled skill directory.
- `.codex-plugin/`, `.claude-plugin/`, `.cursor-plugin/`, `gemini-extension.json`: platform metadata.
