#!/usr/bin/env bash
set -euo pipefail

REPO_SLUG="${WDU_THESIS_SKILL_REPO:-CercaTrovato/wuhan-donghu-thesis-skill}"
BRANCH="${WDU_THESIS_SKILL_BRANCH:-main}"
PLATFORM="${WDU_THESIS_SKILL_PLATFORM:-codex}"
TARGET_SKILLS_DIR="${WDU_THESIS_SKILL_TARGET:-}"
ARCHIVE_PATH="${WDU_THESIS_SKILL_ARCHIVE:-}"

copy_skill_to_skills_dir() {
  local skills_dir="$1"
  local skill_source="$2"

  if [ -z "$skills_dir" ]; then
    echo "Empty skills directory" >&2
    exit 1
  fi

  local destination="$skills_dir/wuhan-donghu-thesis"
  if [ "$(basename "$destination")" != "wuhan-donghu-thesis" ]; then
    echo "Refusing to install to unexpected path: $destination" >&2
    exit 1
  fi

  mkdir -p "$skills_dir"
  rm -rf "$destination"
  cp -R "$skill_source" "$destination"
  echo "Installed wuhan-donghu-thesis -> $destination"
}

install_platform() {
  local platform="$1"
  local skill_source="$2"

  case "$platform" in
    codex)
      copy_skill_to_skills_dir "${AGENTS_HOME:-$HOME/.agents}/skills" "$skill_source"
      copy_skill_to_skills_dir "${CODEX_HOME:-$HOME/.codex}/skills" "$skill_source"
      ;;
    opencode)
      if [ -n "${OPENCODE_SKILLS_HOME:-}" ]; then
        copy_skill_to_skills_dir "$OPENCODE_SKILLS_HOME" "$skill_source"
      elif [ -n "${OPENCODE_CONFIG_DIR:-}" ]; then
        copy_skill_to_skills_dir "$OPENCODE_CONFIG_DIR/skills" "$skill_source"
      else
        copy_skill_to_skills_dir "${XDG_CONFIG_HOME:-$HOME/.config}/opencode/skills" "$skill_source"
      fi
      ;;
    claude)
      copy_skill_to_skills_dir "${CLAUDE_HOME:-$HOME/.claude}/skills" "$skill_source"
      ;;
    cursor)
      copy_skill_to_skills_dir "${CURSOR_HOME:-$HOME/.cursor}/skills" "$skill_source"
      ;;
    all)
      install_platform codex "$skill_source"
      install_platform opencode "$skill_source"
      install_platform claude "$skill_source"
      install_platform cursor "$skill_source"
      ;;
    *)
      echo "Unknown WDU_THESIS_SKILL_PLATFORM '$platform'. Use codex, opencode, claude, cursor, or all." >&2
      exit 1
      ;;
  esac
}

tmp_dir="$(mktemp -d)"
cleanup() {
  rm -rf "$tmp_dir"
}
trap cleanup EXIT

zip_url="https://github.com/$REPO_SLUG/archive/refs/heads/$BRANCH.zip"
zip_path="$tmp_dir/repo.zip"

if [ -n "$ARCHIVE_PATH" ]; then
  echo "Using local archive $ARCHIVE_PATH"
  cp "$ARCHIVE_PATH" "$zip_path"
else
  echo "Downloading $zip_url"
  if command -v curl >/dev/null 2>&1; then
    curl -fsSL "$zip_url" -o "$zip_path"
  elif command -v wget >/dev/null 2>&1; then
    wget -q "$zip_url" -O "$zip_path"
  else
    echo "curl or wget is required" >&2
    exit 1
  fi
fi

if command -v unzip >/dev/null 2>&1; then
  unzip -q "$zip_path" -d "$tmp_dir"
else
  python3 - "$zip_path" "$tmp_dir" <<'PY'
import sys
import zipfile

with zipfile.ZipFile(sys.argv[1]) as zf:
    zf.extractall(sys.argv[2])
PY
fi

repo_root="$(find "$tmp_dir" -mindepth 1 -maxdepth 1 -type d ! -name '__MACOSX' | head -n 1)"
skill_source="$repo_root/skills/wuhan-donghu-thesis"

if [ ! -d "$skill_source" ]; then
  echo "Skill folder not found in archive: $skill_source" >&2
  exit 1
fi

if [ -n "$TARGET_SKILLS_DIR" ]; then
  copy_skill_to_skills_dir "$TARGET_SKILLS_DIR" "$skill_source"
else
  install_platform "$PLATFORM" "$skill_source"
fi
