#!/bin/bash
# SessionStart hook: sprístupní lokálne skilly z tohto repa ako globálne
# skilly (~/.claude/skills/), aby boli dostupné aj v iných/budúcich sessionoch
# bez ohľadu na recykláciu kontajnera. Idempotentné, neinteraktívne.
set -euo pipefail

PROJECT_DIR="${CLAUDE_PROJECT_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)}"
SKILLS_DIR="$HOME/.claude/skills"

mkdir -p "$SKILLS_DIR"

# Zoznam skillov z repa, ktoré chceme mať globálne dostupné.
for skill in vytvor-zos; do
  src="$PROJECT_DIR/$skill"
  dst="$SKILLS_DIR/$skill"
  if [ -d "$src" ]; then
    rm -rf "$dst"
    cp -r "$src" "$dst"
    echo "[session-start] synced skill: $skill -> $dst"
  fi
done

exit 0
