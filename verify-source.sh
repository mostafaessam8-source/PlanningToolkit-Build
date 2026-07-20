#!/usr/bin/env bash
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$root"

required=(
  "PlanningToolkit.sln"
  "src/PlanningToolkit.Core/PlanningToolkit.Core.csproj"
  "src/PlanningToolkit.Infrastructure/PlanningToolkit.Infrastructure.csproj"
  "src/PlanningToolkit.Excel/PlanningToolkit.Excel.csproj"
  "src/PlanningToolkit.Excel/PlanningRibbon.cs"
  "tests/PlanningToolkit.Tests/Program.cs"
  "docs/ROADMAP.md"
  "docs/UserGuide.md"
)

for file in "${required[@]}"; do
  test -s "$file" || { echo "Missing or empty: $file" >&2; exit 1; }
done

if rg -n "TODO|NotImplementedException|PLACEHOLDER" src tests; then
  echo "Unresolved implementation marker found." >&2
  exit 1
fi

test "$(rg -l '<Project Sdk=' src tests -g '*.csproj' | wc -l)" -eq 4
test "$(rg -l 'onAction=' src/PlanningToolkit.Excel/PlanningRibbon.cs | wc -l)" -eq 1
test "$(rg -o 'onAction="[^"]+"' src/PlanningToolkit.Excel/PlanningRibbon.cs | sort -u | wc -l)" -eq 16

echo "Static source verification passed."
