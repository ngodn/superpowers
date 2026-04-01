#!/usr/bin/env bash
# Install or update personal skills from this repo into ~/.claude/skills/
#
# Usage:
#   ./install-personal-skills.sh           # install all custom skills
#   ./install-personal-skills.sh pptx-editing  # install specific skill
#
# Skills are installed to ~/.claude/skills/ which Claude Code loads
# alongside plugin skills. Survives plugin updates.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILLS_SRC="$SCRIPT_DIR/skills"
SKILLS_DST="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"

# Custom skills that are NOT part of the official superpowers plugin.
# Add new personal skills to this list.
PERSONAL_SKILLS=(
    "pptx-editing"
)

install_skill() {
    local name="$1"
    local src="$SKILLS_SRC/$name"
    local dst="$SKILLS_DST/$name"

    if [ ! -d "$src" ]; then
        echo "  skip: $name (not found in $SKILLS_SRC)"
        return 1
    fi

    if [ ! -f "$src/SKILL.md" ]; then
        echo "  skip: $name (no SKILL.md)"
        return 1
    fi

    mkdir -p "$dst"
    cp -r "$src"/* "$dst/"
    echo "  ok: $name -> $dst"
}

echo "Installing personal skills to $SKILLS_DST"
echo ""

if [ $# -gt 0 ]; then
    # Install specific skills
    for name in "$@"; do
        install_skill "$name"
    done
else
    # Install all personal skills
    for name in "${PERSONAL_SKILLS[@]}"; do
        install_skill "$name"
    done
fi

echo ""
echo "Done. Skills available in next Claude Code session."
echo ""
echo "Installed skills:"
ls -1 "$SKILLS_DST" 2>/dev/null | while read d; do
    if [ -f "$SKILLS_DST/$d/SKILL.md" ]; then
        desc=$(grep "^description:" "$SKILLS_DST/$d/SKILL.md" | head -1 | sed 's/description: *//' | tr -d '"')
        printf "  %-20s %s\n" "$d" "$desc"
    fi
done
