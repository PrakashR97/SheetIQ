#!/usr/bin/env bash
# Set or rotate the Basic Auth credentials for SheetIQ (bcrypt hash written to .env only).
# Usage (on the server, from the project dir):
#   scripts/set-basic-auth.sh <username>            # prompts for the password
#   SHEETIQ_PASSWORD=... scripts/set-basic-auth.sh <username>   # non-interactive
set -euo pipefail

user="${1:?usage: $0 <username>}"
cd "$(dirname "$0")/.."

if [[ -n "${SHEETIQ_PASSWORD:-}" ]]; then
  pass="$SHEETIQ_PASSWORD"
else
  read -r -s -p "Password for $user: " pass; echo
  read -r -s -p "Repeat password: " pass2; echo
  [[ "$pass" == "$pass2" ]] || { echo "Passwords differ" >&2; exit 1; }
fi
[[ ${#pass} -ge 12 ]] || { echo "Use at least 12 characters" >&2; exit 1; }

hash="$(printf '%s' "$pass" | docker run --rm -i httpd:2.4-alpine htpasswd -niB "$user")"

touch .env && chmod 600 .env
grep -v '^BASIC_AUTH_USERS=' .env > .env.tmp || true
# single quotes keep the literal $ characters of the hash from being interpolated
printf "BASIC_AUTH_USERS='%s'\n" "$hash" >> .env.tmp
chmod 600 .env.tmp && mv .env.tmp .env

docker compose -f docker-compose.yml -f docker-compose.traefik.yml up -d
echo "Basic Auth updated for user '$user'."
