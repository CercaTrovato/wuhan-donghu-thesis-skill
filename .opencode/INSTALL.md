# Installing Wuhan Donghu Thesis Skill for OpenCode

## Plugin Install

Add the plugin to the `plugin` array in global or project-level `opencode.json`:

```json
{
  "plugin": ["wuhan-donghu-thesis@git+https://github.com/CercaTrovato/wuhan-donghu-thesis-skill.git"]
}
```

Restart OpenCode. The plugin registers the bundled `skills/` directory automatically.

## One-Line Local Skill Install

Windows PowerShell:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$env:WDU_THESIS_SKILL_PLATFORM='opencode'; irm https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.ps1 | iex"
```

macOS / Linux / WSL:

```bash
curl -fsSL https://raw.githubusercontent.com/CercaTrovato/wuhan-donghu-thesis-skill/main/install.sh | WDU_THESIS_SKILL_PLATFORM=opencode bash
```

## Verify

Ask OpenCode to list skills, then load `wuhan-donghu-thesis`.
