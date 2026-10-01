#!/usr/bin/env bash
# Fails when a tracked file holds a personal value.
#
# Generic shapes are built in: a home folder, a workspace mount, a Windows
# user folder, an e-mail address, a manuscript number. Real names, accounts
# and manuscript names come from .personal-patterns.local (untracked, one
# extended regex per line, "#" starts a comment), so none of them is written
# in the repository. Without that file only the generic shapes are checked,
# and the script says so.
#
# Each match is judged on its own: a placeholder (the home folder of "you",
# "user" or a name of one or two letters, an address at example.org) passes,
# and a real value on the same line as a placeholder still fails. scripts/personal-allow.txt lists known
# lines that stay, as fixed text "path:line:" (or any fixed piece of a hit
# line), each with its reason above it.
#
# Output: one "path:line:match" per hit. Exit 0 with "personal check: 0 hits",
# 1 with the hits, 2 outside a git work tree.
set -u
cd "$(git rev-parse --show-toplevel 2>/dev/null)" || { echo "not in a git work tree" >&2; exit 2; }

# Written in two pieces so that this file is no hit of its own patterns.
H='/ho''me'
W='[A-Za-z]:\\+Us''ers\\+'
GENERIC=(
  "$H/[a-z0-9_]+/"
  '/mnt/[a-z]/_work''space'
  "$W"'[^\\ "<>]+'
  '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[a-z]{2,}'
  '[A-Z]{3,5}-D-[0-9]{2}-[0-9]+'
)
# Placeholders, matched against the whole match. A user name of one or two
# letters, "you" or "user" is a placeholder; so is an image file named for its
# pixel density (icon@2x.png).
NAME='([a-z]{1,2}|you|user)'
export PERSONAL_ALLOW="^($H/$NAME/|$W$NAME|.*@example\.(org|com)|noreply@.*|.*@users\.noreply\.github\.com|git@github\.com|.*@[0-9]x\.png)\$"

patterns=("${GENERIC[@]}")
if [ -f .personal-patterns.local ]; then
  while IFS= read -r line || [ -n "$line" ]; do
    case "$line" in '' | '#'*) ;; *) patterns+=("$line") ;; esac
  done < .personal-patterns.local
else
  echo "no .personal-patterns.local: generic shapes only"
fi

allow_lines=$(mktemp)
trap 'rm -f "$allow_lines"' EXIT
if [ -f scripts/personal-allow.txt ]; then
  grep -v '^#' scripts/personal-allow.txt | sed '/^[[:space:]]*$/d' > "$allow_lines"
fi

hits=$(
  for p in "${patterns[@]}"; do
    git grep -I -n -o -E -e "$p" -- . ':!scripts/personal-allow.txt' ':!package-lock.json'
  done |
    awk '{ m = $0; sub(/^[^:]*:[0-9]+:/, "", m); if (m !~ ENVIRON["PERSONAL_ALLOW"]) print }' |
    { if [ -s "$allow_lines" ]; then grep -F -v -f "$allow_lines"; else cat; fi; } |
    sort -u
)
if [ -z "$hits" ]; then
  echo "personal check: 0 hits"
  exit 0
fi
printf '%s\n' "$hits"
exit 1
