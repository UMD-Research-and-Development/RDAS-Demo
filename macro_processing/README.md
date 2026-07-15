# RevRating - InfoReady Review Processor

Cross-platform replacement for the RevRatingUMD VBA macro. Reads any InfoReady Application Reviews export and produces a formatted MacroResult workbook. Works with any competition type (limited submissions, internal funding, fellowships, etc.).

---

## Requirements

- **Python:** 3.8 or later (pre-installed on macOS)
- **Dependencies:**
  - `openpyxl` - reads/writes .xlsx files and applies formatting
  - `pyxlsb` - reads .xlsb files (optional, only needed if input is .xlsb)

### One-Time Setup

Open Terminal and run:

```
pip3 install openpyxl pyxlsb
```

---

## Usage

### Terminal Command

```
python3 rev_rating.py "path/to/InfoReady_Export.xlsx"
```

**Tip:** You can drag the file from Finder directly into the Terminal window after typing `python3 rev_rating.py `.

### Auto-Detection

If you run the script without specifying a file and there is only one .xlsx/.xlsb/.csv file in the current folder, it will auto-detect and use that file. If multiple files are found, it will prompt you to choose.

### Output

The script creates a file named `MacroResult_YYYY-MM-DD_HH-MM.xlsx` in the same folder as the input file. If that folder is read-only (e.g., Google Drive stream), it saves to the current working directory instead.

---

## Input Format

### Accepted File Types

- `.xlsx` (recommended, standard InfoReady export)
- `.xlsb` (legacy macro-enabled binary format)
- `.csv` (fallback)

### Required Sheet

The script looks for a sheet named **Reviewer Details** (case-insensitive partial match). If not found, it looks for any sheet containing "review" in the name, then defaults to the active/first sheet.

### Column Matching

The script uses **flexible column matching** with built-in aliases, so minor naming variations across different InfoReady competition types are handled automatically.

| Field | Required | Recognized Column Names |
|---|---|---|
| Application | Yes | Application, Application Title, Proposal, Proposal Title |
| Reviewer | Yes | Reviewer, Reviewer Name, Panelist, Evaluator |
| Rating | Yes | Rating, Score, Numeric Rating, Numeric Score |
| Applicant | No | Applicant, Applicant Name, PI, PI Name, Principal Investigator |
| Comment Label | No | Comment Label, Label, Rubric Category, Criterion, Category |
| Comments to Applicant | No | Comments to Applicant, Reviewer Comments, Comments, Feedback |
| Administrator Comments | No | Administrator Comments, Admin Comments, Internal Comments |

Additional columns (Co-Applicant, Status, Competition Type, Routing Step, Field Name, Response) are ignored and do not need to be removed.

### Graceful Handling of Missing Columns

- If **Comment Label** is missing: Matrix by Category sheet displays "No rubric categories found"
- If **Comments to Applicant** is missing: Reviewer Comments sheet displays "No reviewer comments found"
- If **Administrator Comments** is missing: Admin Comments sheet displays "No administrator comments found"
- The script never crashes on missing optional columns

---

## Output Structure

The MacroResult workbook contains 5 sheets:

### 1. Matrix Final
- Reviewer-by-application score matrix
- Summary columns: Grand Total, Average, Standard Deviation, Number of Reviews, Normalized Score
- Bottom rows: per-reviewer Grand Total, Average, Std Dev, Count, Deviation From Mean (bias)
- Sorted by Normalized Score (descending)
- Conditional formatting: Standard Deviation > 4 highlighted in light red

### 2. Reviewer Summary
- Per-reviewer metrics: Average, Std Dev, Number of Reviews, Estimated Bias
- Sorted by Estimated Bias (descending, most lenient first)
- Includes bias legend explaining positive vs. negative bias

### 3. Matrix by Category and Applicant
- Cross-tab of scores broken out by rubric category
- Application summary rows (red bold) with reviewer detail sub-rows
- Grand Total column per row
- Adapts automatically to however many rubric categories the competition uses

### 4. Reviewer Administrator Comments
- Admin-only comments organized by reviewer, then by application
- Black header rows for reviewer names, gray rows for application identifiers

### 5. Reviewer Comments
- Applicant-facing comments organized by application, then reviewer
- Each comment prefixed with its rubric category label

---

## VBA Logic Mapping

This table maps each piece of the original VBA macro logic to its Python equivalent for maintenance and future extension.

| Feature | VBA (RevRatingUMD.xlsb) | Python (rev_rating.py) |
|---|---|---|
| Column detection | FindHeaderCol() scans header row by name | col_index() with alias matching |
| Score aggregation | Scripting.Dictionary nested by app/reviewer | defaultdict nested by app/reviewer |
| Population Std Dev | STDEVP() Excel formula | Manual: sqrt(sum((x-mean)^2)/n) |
| Reviewer bias | (SUM - SUMIF(<>, AvgCol)) / COUNT | mean(reviewer_score - proposal_avg) per reviewer |
| Normalized Score | (SUM - SUMIF(<>, BiasRow)) / COUNT | mean(score - bias) across reviewers per proposal |
| Sort by Norm Score | wsOut.Sort descending on Normalized Score | Python sorted() descending, then write back |
| Conditional format | FormatConditions.Add for SD > 4 | CellIsRule(operator="greaterThan", formula=["4"]) |
| Grand Total row | Black fill, white font | Same via PatternFill + Font |
| App key format | Format$(idx,"00") & ". " & Trunc(title,20) & "_" & LastName | f"{idx:02d}. {title[:20]}_{last_name}" |
| Output save | wbOut.SaveAs to source folder | wb.save() to input file directory |

---

## Validation

### Tested Competition Types

| Competition | Apps | Reviewers | Categories | Rows | Result |
|---|---|---|---|---|---|
| Packard Foundation Fellowships FY26 | 10 | 9 | 5 | 150 | Pass |
| RevRatingUMD test dataset | 4 | 3 | 6 | 72 | Pass |

### Spot-Check Results (RevRatingUMD test data)

| Metric | VBA Output | Python Output | Match |
|---|---|---|---|
| Top ranked proposal | 02. Novel Superconductin_Pagli (Norm: 23.67) | 02. Novel Superconductin_Pagli (Norm: 23.67) | Yes |
| Reviewer bias (Reutt-Robey) | +0.33 | +0.33 | Yes |
| Reviewer bias (Briber) | -1.17 | -1.17 | Yes |
| Reviewer bias (Murphy) | +0.83 | +0.83 | Yes |
| SD conditional formatting | Triggers at SD > 4 | Triggers at SD > 4 | Yes |
| Sheet count | 5 | 5 | Yes |
| Application sort order | Descending by Normalized Score | Descending by Normalized Score | Yes |

---

## File Structure

```
processing_script/
  rev_rating.py       # Main script (only file needed to run)
  README.md           # This documentation
```

---

## Troubleshooting

**"openpyxl not installed"** - Run `pip3 install openpyxl pyxlsb` in Terminal.

**"Required columns not found"** - The input file is missing Application, Reviewer, or Rating columns. The script prints a column mapping showing what it found and what is missing. Check that the input file contains the Reviewer Details sheet with standard column headers.

**"No data found"** - The Reviewer Details sheet is empty or the script could not locate it. Check that the sheet name contains "Reviewer" and "Details".

**"No rubric categories found"** - The competition export does not have a Comment Label column. The Matrix by Category sheet will be empty, but all other sheets will still generate normally.

**Output file not appearing** - If the input folder is read-only (e.g., Google Drive stream), the script saves to the current working directory instead. Check Terminal output for the exact save path.
