#!/bin/bash
# Sync project-level skills to user-level ~/.claude/
# Triggered automatically by PostToolUse hook on Edit/Write
#
# Reads tool input from stdin. Only syncs if the edited file is
# under .claude/skills/ to avoid unnecessary work.

input=$(cat)

# Quick check: only sync if the edit touched a skill file
# Match both forward and back slashes (Windows JSON paths use \\)
if ! echo "$input" | grep -qiE '\.claude[/\\]+skills'; then
    exit 0
fi

PROJECT_DIR="C:/Claude/.claude"
USER_DIR="C:/Users/rturpin/.claude"

# Sync skills
for skill_dir in "$PROJECT_DIR"/skills/*/; do
    skill_name=$(basename "$skill_dir")
    mkdir -p "$USER_DIR/skills/$skill_name"
    cp "$skill_dir"SKILL.md "$USER_DIR/skills/$skill_name/SKILL.md" 2>/dev/null
done
