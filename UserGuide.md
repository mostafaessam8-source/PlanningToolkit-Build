# Planning Toolkit — Phase 1 User Guide

## Ribbon Status

The **Planning Toolkit** tab contains active Phase 1 commands and disabled placeholders for upcoming modules. A disabled button is intentional and identifies the phase in which that module will be implemented.

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

## Logs

Use **View Logs** to open the local diagnostics folder. Logs contain command names and technical exceptions; the current phase does not deliberately write workbook cell contents to logs.

## Safety Behavior

Before each worksheet command, Planning Toolkit temporarily disables screen updating, events, alerts and automatic calculation. Their original values are restored after success or failure. Destructive actions require confirmation.

