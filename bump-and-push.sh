#!/bin/bash
set -e

# ----- CONFIG -----
VERSION_FILE=".version"
BRANCH="$(git rev-parse --abbrev-ref HEAD)"
# Tag/Labele prefix (keine doppelte Wiederholung von 'v' später hinzufügen)
TAG_PREFIX="v"
# Maximal sinnvolle Tag-Länge, um OS/Git-Pfadlimits nicht zu reißen
MAX_TAG_LEN=96

# ----- CHECK GIT -----
if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
    echo "❌ Not inside a Git repository!"
    exit 1
fi

if [ "$BRANCH" = "HEAD" ]; then
    echo "❌ You are in a detached HEAD state — aborting."
    exit 1
fi

# ----- BUILD VERSION -----
# Basisteil aus .version lesen (SemVer), sonst 0.0.0
if [ -f "$VERSION_FILE" ] && grep -E '^v?[0-9]+\.[0-9]+\.[0-9]+' "$VERSION_FILE" >/dev/null 2>&1; then
    BASE_VERSION=$(head -n1 "$VERSION_FILE" | sed -E 's/^v?([0-9]+\.[0-9]+\.[0-9]+).*/\1/')
else
    BASE_VERSION="0.0.0"
fi

IFS='.' read -r MAJ MIN PATCH <<< "$BASE_VERSION"
PATCH=$((PATCH + 1))
BASE_VERSION="$MAJ.$MIN.$PATCH"

TS=$(date +%Y%m%d-%H%M)
COMMIT=$(git rev-parse --short HEAD)

NEW_VERSION="${BASE_VERSION}-${TS}-${COMMIT}"
TAG="${TAG_PREFIX}${NEW_VERSION}"

# Safety: Tag nicht zu lang werden lassen
if [ ${#TAG} -gt $MAX_TAG_LEN ]; then
    COMMIT=$(git rev-parse --short=6 HEAD)
    TAG="${TAG_PREFIX}${BASE_VERSION}-${TS}-${COMMIT}"
fi

echo "$NEW_VERSION" > "$VERSION_FILE"
echo "🔢 Version: $BASE_VERSION  ➕  Suffix: $TS-$COMMIT"

# ----- COMMIT -----
git add "$VERSION_FILE"
git commit -m "Bump version to $NEW_VERSION" || true

# ----- TAG -----
echo "🏷️  Creating/Updating tag: $TAG"
git tag -f "$TAG"

# ----- PUSH -----
echo "🚀 Pushing to origin..."
git push origin "$BRANCH"
git push origin "$TAG" --force

echo "✅ Done. Version is now $NEW_VERSION"
