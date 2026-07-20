param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $root "PlanningToolkit.sln"
$artifacts = Join-Path $root "artifacts\phase1"
$excelOutput = Join-Path $root "src\PlanningToolkit.Excel\bin\$Configuration\net8.0-windows"

Write-Host "Planning Toolkit Phase 1 build"
dotnet --version
dotnet restore $solution
dotnet build $solution --configuration $Configuration -p:Platform=x64 --no-restore
dotnet run --project (Join-Path $root "tests\PlanningToolkit.Tests\PlanningToolkit.Tests.csproj") --configuration $Configuration --no-build

if (Test-Path $artifacts) {
    Remove-Item $artifacts -Recurse -Force
}
New-Item -ItemType Directory -Path $artifacts | Out-Null

Get-ChildItem $excelOutput -File | Where-Object {
    $_.Extension -in ".xll", ".dll", ".json", ".config", ".dna"
} | Copy-Item -Destination $artifacts

$docsOutput = Join-Path $artifacts "docs"
New-Item -ItemType Directory -Path $docsOutput | Out-Null
Copy-Item (Join-Path $root "README.md") $artifacts
Copy-Item (Join-Path $root "docs\UserGuide.md") $docsOutput
Copy-Item (Join-Path $root "docs\ROADMAP.md") $docsOutput

$xllFiles = Get-ChildItem $artifacts -Filter "*.xll"
if ($xllFiles.Count -eq 0) {
    throw "Build completed but no Excel-DNA XLL was produced. Review the ExcelDna.AddIn build output."
}

Write-Host "Phase 1 build and tests completed."
Write-Host "Artifacts: $artifacts"
$xllFiles | ForEach-Object { Write-Host "XLL: $($_.FullName)" }

