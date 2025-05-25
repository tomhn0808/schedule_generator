# Schedule Generator Documentation

This script generates a customizable schedule of time slots from a CSV input file and outputs the result as a new CSV. It supports grouping, variable slot durations, break periods, and handles both French and English headers.

---

## Overview

* **Reads** an input CSV (Windows-1252 encoding) containing columns `Nom`/`Name` and `Groupe` or `Group`.
* **Filters** entries based on one or more groups (A–E or `/` for all).
* **Shuffles** entries, assigns time slots of fixed duration.
* **Skips** optional break periods.
* **Writes** an output CSV (UTF-8 with BOM) containing only `Nom`/`Name`, `Prénom`/`FirstName`, and `Horaire`/`Schedule`. Optionally, writes the output in PDF or/and HTML format

> ⚠️ **Warning:** Ensure the input CSV file is **closed** (not open in Excel or another application) when running this script to prevent file access errors.

---

## Parameters

| Name               | Type             | Required | Default           | Description                                                                                                                    |
| ------------------ | ---------------- | -------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `inputFile`        | `FileInfo`       | Yes      | N/A               | Path to the input CSV file. Must include at least one header `Nom` or `Name`, and a `Groupe` or `Group` column.                           |
| `group`            | `string[]`       | Yes      | N/A               | One or more group letters (`A`–`E`) or `/` to accept all A–E groups. Case-insensitive. Examples: `-group A,C,E` or `-group /`. |
| `outputFile`       | `FileInfo`       | No       | `.\schedule.csv` | Path to the output CSV. Will contain filtered columns plus the assigned time slot (`Horaire`).                                 |
| `slotDuration`     | `int`            | No       | `120`             | Duration of each time slot in minutes. Valid range: 1–240.                                                                     |
| `nbStudentPerSlot` | `int`            | No       | `4`               | Number of entries per time slot.                                                                                               |
| `scheduleStart`    | `string (HH:mm)` | No       | `"08:30"`         | Start time of the first slot, in `HH:mm` format.                                                                               |
| `scheduleEnd`      | `string (HH:mm)` | No       | `"17:00"`         | End time of the last slot, in `HH:mm` format.                                                                                  |
| `breakTimes`       | `string[]`       | No       | `@[""]`           | Optional list of break periods in `HH:mm-HH:mm` format, comma-separated. Example: `-breakTimes "10:00-10:30","12:00-13:00"`.   |
| `htmlReport`       | `Switch`         | No       | `false`           | Generates an HTML report of the schedule using the `PSHTML` module. Requires user input for the report title.                  |
| `pdfReport`        | `Switch`         | No       | `false`           | Generates a PDF report using the "Microsoft Print to PDF" printer. User is prompted to name the PDF file before saving.        |
---

## Description of Behavior

1. **Input reading**
   Reads the entire file as Windows-1252 to preserve French accented characters, then splits into lines.

2. **Header detection & cleanup**
   Locates the header line matching `Name` or `Nom`, discards any leading/trailing empty or delimiter-only lines, normalizes delimiters, and writes a temporary UTF-8 CSV.

3. **Import & filter**
   Imports the cleaned CSV, excludes rows where `Nom`/`Name` equals the header text, then filters by the `groupe` field:

   * Extracts the letter `A`–`E` from the `groupe` value.
   * If `group` parameter contains `/`, accepts all valid letters.
   * Otherwise only keeps entries whose letter appears in the `group` list.
   * Writes a red warning for any excluded row.

4. **Scheduling loop**

   * Converts `slotDuration`, `scheduleStart`, and `scheduleEnd` into `TimeSpan`/`DateTime`.
   * Parses `breakTimes` into start/end `DateTime` intervals.
   * Iterates, assigning each entry a `Horaire` of the current slot unless it falls inside a break—skipping and jumping to end of break if so.
   * Advances the current time by `slotDuration` after each group of entries.

5. **Output**

   * Selects only specific properties (`Name`,`FirstName`,`Schedule`). Customisation for the output properties is in coding process.
   * Exports the final schedule as a UTF-8 BOM CSV at `outputFile`.
   * Provides an option to generate an **HTML report** using the `PSHTML` module, where users can define a custom title.
   * Allows the **PDF report** generation via the "Microsoft Print to PDF" printer, requiring user input for file naming.
   * Exits with an error code if writing fails.

---
## Requirements

This script requires the **PSHTML** module to generate an HTML report. If it is not installed, you can install it using the following command:
```powershell
Install-Module PSHTML -Scope CurrentUser
```
> This command requires Administrator privileges.

---
## Examples

```powershell
.\schedule.ps1 `
  -inputFile .\data.csv `
  -group A,C,E `
  -slotDuration 90 `
  -nbStudentPerSlot 3 `
  -scheduleStart 09:00 `
  -scheduleEnd 16:30 `
  -breakTimes "12:00-12:30","15:00-15:15"
```

```powershell
.\schedule.ps1 `
  -inputFile .\data.csv `
  -group / `
  -breakTimes "13:00-14:00"
```

---

## Notes

* The script **throws** a terminating error if a row’s `groupe` letter is invalid or not permitted.
* Handles CSVs with headers in French (`Nom`, `Prénom`) or English (`Name`, `FirstName`).
* Break periods are strictly skipped—partial overlaps jump to the break end before assigning the next slot.
* For more details, use `Get-Help .\schedule.ps1 -Full`.
* Example files are available on the Github `dataFrench.csv` and `dataEnglish.csv`.

