#!/bin/sh
# Puts the ComOffice .NET package of a release of ComOffice for Windows and Linux in packages/, where nuget.config
# looks for it, in place of any package there before.
#
#   scripts/fetch-package.sh            the latest release
#   scripts/fetch-package.sh 1.6.0      that release
#
# GH_TOKEN or GITHUB_TOKEN, when set, is sent to GitHub's API, so a CI runner sharing an address stays under the
# anonymous rate limit.
set -eu
cd "$(dirname "$0")/.."
repo=ComOfficeKit/ComOfficeKit-Windows
version="${1:-}"
if [ -z "$version" ]; then
  token="${GH_TOKEN:-${GITHUB_TOKEN:-}}"
  if [ -n "$token" ]; then
    latest=$(curl -fsSL -H "Authorization: Bearer $token" "https://api.github.com/repos/$repo/releases/latest")
  else
    latest=$(curl -fsSL "https://api.github.com/repos/$repo/releases/latest")
  fi
  version=$(printf '%s' "$latest" | sed -n 's/.*"tag_name": *"v\([^"]*\)".*/\1/p' | head -n 1)
  [ -n "$version" ] || { echo "could not read the latest release of $repo" >&2; exit 1; }
fi
mkdir -p packages
file="packages/ComOfficeKit.$version.nupkg"
curl -fL -o "$file.part" "https://github.com/$repo/releases/download/v$version/ComOfficeKit.$version.nupkg"
rm -f packages/ComOfficeKit.*.nupkg
mv "$file.part" "$file"
echo "$file"
