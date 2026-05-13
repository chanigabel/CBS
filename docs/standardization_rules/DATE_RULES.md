````md
# DATE_RULES.md

# Date Rules & Logic

## Purpose

The date engine receives messy date values from Excel workbooks, parses them using deterministic rules, generates corrected fields, and writes visible statuses whenever recovery, completion, ambiguity, invalidity, or special parsing occurs.

The engine must:
- preserve original values
- avoid silent incorrect corrections
- avoid inventing dates without a clear rule
- remain deterministic
- survive malformed input safely
- provide visible status for suspicious behavior

---

# Core Principles

## Original values are immutable

Original imported values must never be overwritten.

Examples:
- birth_date
- entry_date
- birth_day
- birth_month
- birth_year
- entry_day
- entry_month
- entry_year

Original values are preserved for:
- auditability
- debugging
- export traceability
- comparison against corrected values

---

## Corrections go only into corrected fields

All parsed or corrected values must be written only into corrected fields.

Examples:
- birth_year_corrected
- birth_month_corrected
- birth_day_corrected
- entry_year_corrected
- entry_month_corrected
- entry_day_corrected

Optional display-only fields if implemented:
- birth_date_corrected
- entry_date_corrected

---

# Corrected Date Field Contract

The active export contract is component-based:

- birth_year_corrected
- birth_month_corrected
- birth_day_corrected
- entry_year_corrected
- entry_month_corrected
- entry_day_corrected

Combined fields such as:
- birth_date_corrected
- entry_date_corrected

are optional display fields only if explicitly implemented.

---

## Suspicious behavior requires visible status

If the system:
- defaults missing values
- performs fallback parsing
- recovers dates from malformed values
- recovers dates from split columns
- parses serial dates
- removes invalid trailing text
- detects ambiguity
- detects invalidity

then a visible status must be written.

---

## Status texts should be Hebrew

User-facing statuses should remain Hebrew.

---

# Status Severity

Date statuses are user-visible explanations. They may represent:
- blocking error
- warning
- recovery note
- informational parse note

A non-empty status does not automatically mean the corrected value is invalid.

---

## Invalid input must not crash the system

Malformed or invalid date values must never crash:
- the parser
- the pipeline
- export logic
- UI rendering

Instead:
- corrected values may remain blank
- visible status should explain the problem

---

# Main Date Fields

## Source fields

```text
birth_date
entry_date
birth_day
birth_month
birth_year
entry_day
entry_month
entry_year
````

## Corrected fields

```text
birth_year_corrected
birth_month_corrected
birth_day_corrected
entry_year_corrected
entry_month_corrected
entry_day_corrected
```

Optional display fields:

```text
birth_date_corrected
entry_date_corrected
```

## Status fields

```text
birth_date_status
entry_date_status
```

---

# Reference Year Rules

## Two-digit year expansion

Two-digit years are resolved using the processing/reference year.

Example with reference year 2026:

```text
25 -> 2025
26 -> 2026
27 -> 1927
99 -> 1999
00 -> 2000
```

Rule:

```text
if YY <= last two digits of reference year:
    resolve as 20YY
else:
    resolve as 19YY
```

---

# Important Exception: Split Year 0

In split year columns:

```text
0
"0"
```

must NOT be treated as:

* 2000
* a valid two-digit year
* a valid numeric year

Instead:

* treat as empty/missing
* corrected year remains blank

Examples:

```text
birth_year = 0 -> birth_year_corrected = ""
entry_year = "0" -> entry_year_corrected = ""
```

---

# Empty Values

The following values are considered empty:

```text
None
""
"   "
0 in split year column
"0" in split year column
```

Expected behavior:

* do not create fake dates
* do not auto-complete to 2000
* corrected fields remain blank where appropriate
* write missing/empty status if required

Possible statuses:

```python
STATUS_EMPTY_CELL = "תא ריק"
STATUS_MISSING_YEAR = "חסר שנה"
STATUS_MISSING_MONTH = "חסר חודש"
STATUS_MISSING_DAY = "חסר יום"
STATUS_MISSING_MONTH_DAY = "חסר חודש ויום"
```

---

# Single-Field Date Parsing

The engine should support formats such as:

```text
01/02/2024
1/2/24
01.02.2024
01-02-2024
2024-02-01
01//02//2024
01/02/2024abc
```

---

# Repeated Separators

Values like:

```text
01//02//2024
```

must normalize repeated separators before parsing.

Equivalent normalized value:

```text
01/02/2024
```

Examples:

```text
01..02..2024 -> 01.02.2024
01--02--2024 -> 01-02-2024
```

---

# Trailing Text Cleanup

Values like:

```text
01/02/2024abc
```

should:

* recover the valid date portion
* populate corrected fields
* write visible status indicating extra text was ignored

The system must not:

* silently ignore the issue
* crash
* treat the value as perfectly clean

---

# Two-Part Dates

Examples:

```text
1/2
01/02
```

These contain only:

* day/month
  or
* month/day

with no year.

Expected behavior:

* the engine may default the year from the reference year
* but the result must NOT appear fully clean
* a visible status is required

Example:

```python
STATUS_MISSING_YEAR_DEFAULTED = "שנה חסרה והושלמה"
```

---

# Compact Numeric Dates

## Important Distinction

Do NOT confuse:

* compact numeric dates

with:

* Excel serial dates

Compact numeric dates are numeric strings whose digits encode a date structure.

Examples:

```text
01022024
010224
112024
2024
1124
```

Excel serial dates are Excel-internal numeric date values.

Examples:

```text
36525
38353
45657
```

---

# Numeric Safety Rule

Plain numeric Excel cells must not be treated as Excel serial dates unless extraction metadata confirms the source cell was date-formatted.

Examples:

```text
36525 with date metadata -> parsed as Excel serial
36525 as General numeric -> rejected
2024 as General numeric -> handled only by compact/year-only rules
1234567 -> rejected
```

---

# 8-Digit Compact Parsing

## First attempt

Parse as:

```text
DD MM YYYY
```

Example:

```text
01022024 -> 01/02/2024
```

---

## Second attempt

If DD/MM/YYYY is invalid:

Try:

```text
MM DD YYYY
```

Example:

```text
12312024 -> 12/31/2024
```

---

## If both attempts fail

Do NOT:

* invent another interpretation
* convert into Excel serial

Return:

* invalid-date status
* invalid-day/month/year status
* invalid-format status

---

# 6-Digit Compact Parsing

## First attempt

Parse as:

```text
DD MM YY
```

Example:

```text
010224 -> 01/02/2024
```

---

## Second attempt

If invalid:

```text
MM DD YY
```

Example:

```text
123124 -> 12/31/2024
```

---

## Third attempt

If still invalid:

```text
D M YYYY
```

Example:

```text
112024 -> 01/01/2024
```

---

## If all attempts fail

* do not invent dates
* do not convert into serial
* return deterministic invalid status

---

# 4-Digit Compact Parsing

## First attempt: standalone year

Example:

```text
2024
```

Behavior:

```text
year_corrected = 2024
month_corrected = ""
day_corrected = ""
```

Status:

```python
STATUS_MISSING_MONTH_DAY = "חסר חודש ויום"
```

---

## Second attempt

Parse as:

```text
D M YY
```

Example:

```text
1124 -> 01/01/2024
```

---

## Third attempt

Parse as:

```text
M D YY
```

---

## If all attempts fail

* do not invent dates
* return invalid status

---

# 5/7 Digit Numeric Values

Five-digit and seven-digit numeric values are ambiguous.

Examples:

```text
10224
1234567
```

Default behavior:

* do not parse automatically
* preserve original value
* leave corrected fields blank
* write visible ambiguity status

Optional future heuristic mode:

* may parse only if explicitly enabled
* only if exactly one valid interpretation exists
* must write visible recovery status

---

# Excel Serial Dates

## What is an Excel serial date

Excel sometimes stores dates as serial integers.

Examples:

```text
36525
38353
45657
```

---

## Recognized serial date

If the value is recognized as a valid serial date AND metadata confirms the source cell was date-formatted:

* convert to date
* populate corrected fields
* write visible status

Status:

```python
STATUS_EXCEL_SERIAL_PARSED = "פורק מתאריך סידורי"
```

---

## Unrecognized serial candidate

If a number is NOT recognized as a valid serial date:

* do not silently convert it
* do not present it as clean date
* write visible status

Status:

```python
STATUS_EXCEL_SERIAL_NOT_RECOGNIZED = "מספר לא הוכר כתאריך"
```

---

# Split Date Columns

The engine supports split date columns:

```text
day | month | year
```

Examples:

```text
birth_day
birth_month
birth_year
entry_day
entry_month
entry_year
```

---

# Split Year 0

Values:

```text
0
"0"
```

mean:

* empty
* missing

NOT:

* 2000
* valid year

Behavior:

```text
year_corrected = ""
```

---

# Full Date Inside Split Column

Sometimes a full date appears inside one split column.

Examples:

```text
birth_day = 11.06.1997
birth_month = ""
birth_year = ""
```

---

## Expected behavior

If remaining split fields are sufficiently empty and non-conflicting:

* recover the full date
* populate corrected fields
* write visible source status

Statuses:

```python
STATUS_SPLIT_FULL_DATE_FROM_DAY = "תאריך מלא פורק מעמודת יום"
STATUS_SPLIT_FULL_DATE_FROM_MONTH = "תאריך מלא פורק מעמודת חודש"
STATUS_SPLIT_FULL_DATE_FROM_YEAR = "תאריך מלא פורק מעמודת שנה"
```

---

# Split-Date Conflicts

Example:

```text
day = 11.06.1997
month = 5
year = 1998
```

Expected behavior:

* do not silently choose interpretation
* do not invent a clean date
* write visible conflict status

Status:

```python
STATUS_SPLIT_FULL_DATE_CONFLICT = "ערכים סותרים בעמודות תאריך מפוצלות"
```

---

# Invalid Calendar Dates

Examples:

```text
31/02/2020
29/02/2023
32/01/2020
00/12/2020
01/13/2020
```

Expected behavior:

* do not crash
* partial corrected components may remain
* invalid dates must never appear clean
* visible invalid status is required

---

# Birth Date Rules

## Future birth date

```python
STATUS_FUTURE_BIRTH = "תאריך לידה עתידי"
```

---

## Birth year before 1906

```python
STATUS_BEFORE_1906 = "שנה לפני 1906"
```

---

# Entry Date Rules

## Future entry date

```python
STATUS_FUTURE_ENTRY = "תאריך כניסה עתידי"
```

---

## Late entry date

```python
STATUS_LATE_ENTRY = "תאריך כניסה מאוחר מהתאריך שנקבע"
```

---

# Approved Existing Behaviors

## Majority correction

Majority correction currently remains in the system.

Do not:

* remove it
* redesign it

unless explicitly requested.

---

## Workbook-level date-format detection

Workbook-level DD/MM vs MM/DD pattern detection currently remains allowed.

Do not:

* remove it
* redesign it

unless explicitly requested.

---

# Export Rules

During export:

* corrected fields should be preferred
* original values must not overwrite corrected results
* suspicious dates must not appear as clean dates without status visibility
* invalid dates must not silently reconstruct into valid dates
* partial corrected values + status should remain traceable

---

# Existing Statuses

```python
STATUS_EMPTY_CELL = "תא ריק"
STATUS_INVALID_DATE_VALUE = "ערך תאריך לא תקין"
STATUS_UNPARSEABLE = "תוכן לא ניתן לפריקה"
STATUS_INVALID_DAY = "יום לא תקין"
STATUS_INVALID_MONTH = "חודש לא תקין"
STATUS_INVALID_YEAR = "שנה לא תקינה"
STATUS_DATE_NOT_EXISTS = "תאריך לא קיים"
STATUS_BEFORE_1906 = "שנה לפני 1906"
STATUS_LATE_ENTRY = "תאריך כניסה מאוחר מהתאריך שנקבע"
STATUS_FUTURE_BIRTH = "תאריך לידה עתידי"
STATUS_FUTURE_ENTRY = "תאריך כניסה עתידי"
STATUS_MISSING_MONTH_DAY = "חסר חודש ויום"
STATUS_INVALID_LENGTH = "אורך תאריך לא תקין"
STATUS_UNCLEAR_DATE = "תאריך לא ברור"
STATUS_INVALID_FORMAT = "פורמט תאריך לא תקין"
STATUS_UNRECOGNIZED_FORMAT = "פורמט תאריך לא מזוהה"
STATUS_NO_SEPARATOR = "אין מפריד בתאריך"
STATUS_MISSING_YEAR = "חסר שנה"
STATUS_MISSING_MONTH = "חסר חודש"
STATUS_MISSING_DAY = "חסר יום"
STATUS_MISSING_YEAR_DEFAULTED = "שנה חסרה והושלמה"
STATUS_EXCEL_SERIAL_PARSED = "פורק מתאריך סידורי"
STATUS_EXCEL_SERIAL_NOT_RECOGNIZED = "מספר לא הוכר כתאריך"
STATUS_SPLIT_FULL_DATE_CONFLICT = "ערכים סותרים בעמודות תאריך מפוצלות"
STATUS_SPLIT_FULL_DATE_FROM_DAY = "תאריך מלא פורק מעמודת יום"
STATUS_SPLIT_FULL_DATE_FROM_MONTH = "תאריך מלא פורק מעמודת חודש"
STATUS_SPLIT_FULL_DATE_FROM_YEAR = "תאריך מלא פורק מעמודת שנה"
```

---

# Final Principle

The engine should be:

```text
tolerant during parsing
strict during status reporting
deterministic
non-destructive
stable
auditable
```

```
```
