# Planning Toolkit — Phase 6 User Guide

## Ribbon Status

The **Planning Toolkit** tab contains XER import, validation, safe export, WBS/Gantt, comparison, PMS and Excel data tools.

## Import, Edit and Export XER

1. Click **Import XER** and select the original Primavera XER file.
2. Save the generated Excel workbook before making changes.
3. Editable XER headers and cells are highlighted in pale yellow/blue. Gray headers are read-only identifiers, references or unsupported fields.
4. Edit permitted values without inserting/deleting rows or columns and without renaming worksheets or XER headers.
5. Click **Export XER**, choose the output `.xer` name, then wait for the success message.

Version 0.6.0 permits common project, WBS, activity, relationship, resource assignment, resource, calendar, activity-code and UDF value edits. IDs and references remain read-only. Adding or deleting XER records is intentionally blocked in this version.

Before writing the final file, Planning Toolkit checks worksheet structure, read-only values, date and numeric formats, percentage limits, unique IDs, project/WBS/activity/resource/calendar references and self-predecessors. It creates an automatic source backup, writes to a temporary file, re-imports the generated XER and verifies that every table, field and value survived the round trip exactly. Export is stopped when a critical check fails.

Keep the original XER file available at the path shown in **XER Summary**. If it was moved, Export XER asks you to select it again. Unknown and unsupported XER tables are preserved from that original file.

## Text & Data Commands

### Fill Down

Select two or more rows in one continuous range. The first selected row is filled down using Excel's native Fill Down behavior, including relative formula adjustment.

### Trim

Removes leading and trailing whitespace and reduces repeated spaces or tabs inside text to a single space. Formulas and non-text values are preserved.

### Clean

Removes non-printing control characters while retaining tabs and line breaks. Formulas and non-text values are preserved.

### Change Case

Converts selected text to upper, lower, proper or sentence case. Formula cells are preserved.

### Text to Date

Recognizes common English, international and Arabic date strings and replaces successfully parsed text with real Excel date serials. The number format comes from Settings. Unrecognized values remain unchanged.

## Range Utilities

### Split Text

Select one column, provide a delimiter, and confirm the destination width. Use `\t` to split on a tab. This operation can write into columns to the right of the selected column.

### Merge Text

Select two or more columns and provide a separator. Results are written to the first column immediately to the right of the selection after confirmation.

### Unique Values

Creates a new worksheet with non-empty distinct values from the selection. Matching is case-insensitive.

### Remove Blank Rows

Finds rows that are blank across the selected columns. After confirmation, the command deletes the corresponding complete worksheet rows.

## Settings

Settings include date format, currency, working hours, schedule thresholds, planned future Gantt/WBS colors, output folder, calculation behavior and logging level. Settings are validated before being saved.

## PMS Dashboard and Look Ahead

Select the Baseline XER and then the Update XER. Planning Toolkit asks for the Look Ahead duration in whole weeks from 1 to 52; the default is 2 weeks.

The generated Look Ahead, Critical Path and WBS Progress worksheets use the XER WBS hierarchy. Bold WBS subtotal rows roll up all descendant activities. Use Excel's outline controls to expand or collapse each WBS group and show activity-level detail.

The WBS Progress subtotals contain activity count, baseline cost, planned and actual percentages, variance, delayed count, critical count and maximum finish variance.

Planned progress uses the Update data date against each activity's Baseline target start and finish dates. Project and WBS roll-ups use Baseline budgeted cost from TASKRSRC and PROJCOST. If no valid Baseline cost exists, the report explicitly identifies a duration-weighted fallback.

## Logs

Use **View Logs** to open the local diagnostics folder. Logs contain command names and technical exceptions; the current phase does not deliberately write workbook cell contents to logs.

## Safety Behavior

Before each worksheet command, Planning Toolkit temporarily disables screen updating, events, alerts and automatic calculation. Their original values are restored after success or failure. Destructive actions require confirmation.
