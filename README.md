# Planning Toolkit

Planning Toolkit is an original, independently implemented 64-bit Microsoft Excel add-in foundation for planning and project-controls workflows. The project does not contain or depend on proprietary Plannex code, branding, icons, licensing logic, or decompiled components.

Version 0.6.0 provides:

- Excel-DNA 64-bit add-in project targeting .NET 8 for Windows.
- Original **Planning Toolkit** Ribbon tab.
- Safe Excel state capture and restoration around every command.
- Daily local logging and a log-folder command.
- Persistent JSON settings with validation and atomic replacement.
- Fill Down, Trim, Clean, case conversion and Text-to-Date commands.
- Split Text, Merge Text, Unique Values and Remove Blank Rows commands.
- A dependency-free automated test runner for the Core and Infrastructure projects.
- XER import and structural validation without requiring Primavera installation.
- Hierarchical WBS schedules, Gantt charts, Baseline/Update comparison and delayed-activity reporting.
- Cost-weighted PMS dashboards, S-curves, configurable Look Ahead, Critical Path and WBS Progress groups/subtotals.
- Safe XER editing/export with editable-field highlighting, read-only ID protection, automatic backups and exact round-trip verification.

The public feature descriptions used to define the wider product roadmap are documented in [ROADMAP.md](docs/ROADMAP.md). The implementation itself is original.

## Requirements

- Windows 10 or Windows 11, 64-bit.
- Microsoft Excel 2016 or later, 64-bit.
- Visual Studio 2022 17.8+ with the **.NET desktop development** workload, or .NET 8 SDK.
- .NET 8 Desktop Runtime x64 on the computer that will load the add-in in Excel.
- Internet access during the first restore of the ExcelDna.AddIn NuGet package.

## Build

From PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\build-phase1.ps1
```

The script restores packages, builds Release x64 and runs the automated tests. GitHub Actions packages the generated 64-bit `.xll`.

Manual commands:

```powershell
dotnet restore .\PlanningToolkit.sln
dotnet build .\PlanningToolkit.sln --configuration Release -p:Platform=x64
dotnet run --project .\tests\PlanningToolkit.Tests\PlanningToolkit.Tests.csproj --configuration Release
```

## Load in Excel

1. Close all Excel windows after building.
2. Open Excel and go to **File → Options → Add-ins**.
3. At the bottom select **Excel Add-ins**, then click **Go**.
4. Click **Browse** and select the generated 64-bit `.xll` from `artifacts\phase1`.
5. Confirm that the **Planning Toolkit** tab appears.

If Windows blocks the file, right-click the `.xll`, open **Properties**, select **Unblock**, and retry.

## Project Structure

```text
src/PlanningToolkit.Core            Pure settings and text/date logic
src/PlanningToolkit.Infrastructure  JSON settings, paths and file logging
src/PlanningToolkit.Excel           Excel-DNA add-in, Ribbon, UI and Excel commands
tests/PlanningToolkit.Tests         Dependency-free automated test runner
docs                                Roadmap and user guide
scripts                             Windows build and packaging commands
```

## Current Limitations

- XER export v0.6.0 edits existing rows only. Adding/deleting XER records and changing IDs or references is blocked for safety.
- The original source XER must remain available so unsupported fields and tables can be preserved exactly.
- Excel integration requires Windows and cannot be executed inside a Linux CI container without Excel.
- Split Text and Merge Text request confirmation because they write outside part of the original selection.
- Remove Blank Rows deletes complete worksheet rows after explicit confirmation.

## Technology References

- [Excel-DNA official documentation](https://excel-dna.github.io/)
- [Excel-DNA Ribbon documentation](https://excel-dna.github.io/docs/guides-basic/customizing-ribbons/)
- [ExcelDna.AddIn package](https://www.nuget.org/packages/ExcelDna.AddIn/)
- [Excel-DNA .NET runtime support](https://excel-dna.net/docs/guides-basic/dotnet-runtime-support/)
