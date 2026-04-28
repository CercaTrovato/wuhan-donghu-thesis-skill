# Installing Wuhan Donghu Thesis Skill for Codex

## One-Line Install

Windows PowerShell:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.ps1 | iex"
```

macOS / Linux / WSL:

```bash
curl -fsSL https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.sh | bash
```

The installer copies the skill to both:

- `~/.agents/skills/wuhan-donghu-thesis`
- `~/.codex/skills/wuhan-donghu-thesis`

## Plugin Package

This repository also includes `.codex-plugin/plugin.json`, so it can be packaged as a Codex plugin or added to a plugin marketplace later.
