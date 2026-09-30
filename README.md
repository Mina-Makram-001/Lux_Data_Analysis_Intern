# Lux Insurance Data Pipeline

## 1. Project Overview

A modular Python pipeline for processing, cleaning, validating, and reconciling insurance claims data from Excel sources.

### Pipeline

```text
Excel Sources
     ↓
Data Loading
     ↓
Column / Schema Standardization
     ↓
LOB & Status Cleaning
     ↓
Date Cleaning
     ↓
Data Quality & Business-Rule Validation
     ↓
LOB-Level Financial Aggregation
     ↓
Financial Reconciliation
     ↓
Excel Reports
```

### Main Outputs

- `Tests_Report.xlsx` — consolidated data-quality and validation results.
- `Financials_Test_Report.xlsx` — calculated financial aggregates and reconciliation differences.

---

# 2. Project Structure

```text
Lux_Project_Intern/
│
├── Data/
│   ├── 1_DATA.xlsx
│   ├── 2_LIST_OF_COLUMNS.xlsx
│   ├── 3_LOB.xlsx
│   ├── 4_TESTS.xlsx
│   └── 5_FINANCIALS.xlsx
│
├── Reporting_Functions/
│   ├── COLUMN_NAMES_CLEAN.py
│   ├── DATA_LOAD.py
│   ├── DATA_TESTS.py
│   └── LOB_AND_STATUS_CLEAN.py
│
├── Report_Exported/
│   ├── Tests_Report.xlsx
│   └── Financials_Test_Report.xlsx
│
└── Data_Pipeline.ipynb
```

The notebook is the orchestration layer; reusable transformation and validation logic is kept in separate Python modules.

---

# 3. `Data_Pipeline.ipynb`

The notebook coordinates the complete workflow.

## 3.1 Load the Sources

```python
from Reporting_Functions import *
import pandas as pd
import numpy as np
from rapidfuzz import process, fuzz

df_1_DATA = load_excel(r"...\Data\1_DATA.xlsx")
df_2_LIST_OF_COLUMNS = load_excel(r"...\Data\2_LIST_OF_COLUMNS.xlsx")
df_3_LOB = load_excel(r"...\Data\3_LOB.xlsx")
df_4_TESTS = load_excel(r"...\Data\4_TESTS.xlsx")
df_5_FINANCIALS = pd.read_excel(
    r"...\Data\5_FINANCIALS.xlsx",
    header=[0, 1]
)
```

Five sources are loaded with different roles:

| Source | Role |
|---|---|
| `1_DATA.xlsx` | Main claims data |
| `2_LIST_OF_COLUMNS.xlsx` | Target schema |
| `3_LOB.xlsx` | Approved LOB and status values |
| `4_TESTS.xlsx` | Test/reference input |
| `5_FINANCIALS.xlsx` | Financial reference totals |

The important design choice is that the reference files are used to drive the cleaning and validation instead of hard-coding all expected values inside the notebook.

---

# 4. `COLUMN_NAMES_CLEAN.py`

## Purpose

Standardize inconsistent source column names against the expected schema.

The module combines **explicit business mappings** with **fuzzy matching**.

## 4.1 Known Synonyms

```python
SYNONYMS = {
    "AS_OF": "LA_AS_AT_DATE",
    "ACCINDENT_DATE": "LA_LOSS_DATE",
    "PAID_DATE": "LA_PAYMENT_DATE",
    "MAJOR_LINE_OF_BUISNESS": "LA_LOB",
}
```

Known variations are handled deterministically before fuzzy matching. This is safer for business-specific names because a known mapping does not depend on a similarity score.

## 4.2 Schema Matching

The main function is:

```python
def clean_and_align_columns(df, to_list, synonyms=SYNONYMS):
```

The core matching logic:

```python
best_match, score, match_key = process.extractOne(
    col,
    pool,
    scorer=fuzz.WRatio
)

new_columns.append(match_key)
del pool[match_key]
```

### Logic

```text
Incoming column
      ↓
Normalize name
      ↓
Known synonym?
   ↙       ↘
 Yes        No
  ↓          ↓
Direct      RapidFuzz
mapping     WRatio
   ↘       ↙
 Standard target
      ↓
Remove target from pool
```

Removing the matched target from `pool` is important: it prevents multiple source columns from being assigned to the same standardized field.

### Why this matters

The pipeline can tolerate minor naming differences and spelling inconsistencies while still producing a consistent internal schema.

---

# 5. `LOB_AND_STATUS_CLEAN.py`

## Purpose

Standardize categorical values using controlled reference values from `3_LOB.xlsx`.

The same approach is applied to:

- `LA_LOB`
- `LA_STATUS`

## 5.1 Normalization

```python
def normalize(text):
    return str(text).strip().upper().replace("_", " ")
```

This creates a common comparison format.

For example:

```text
motor_comprehensive
Motor Comprehensive
 MOTOR_COMPREHENSIVE
```

become comparable representations.

## 5.2 Fuzzy Reference Matching

```python
def fuzzy_map_column(series, choices, threshold=75):
```

The function builds a mapping only once per unique source value:

```python
for val in series.dropna().unique():
    result = process.extractOne(
        normalize(val),
        normalized_choices,
        scorer=fuzz.WRatio
    )

    if result and result[1] >= threshold:
        _, score, idx = result
        mapping[val] = choices[idx]
    else:
        mapping[val] = val
```

### Important behavior

- Matching is performed against approved reference values.
- `WRatio` handles spelling and formatting differences.
- `threshold=75` controls the minimum similarity.
- Values below the threshold are retained rather than silently forced into an incorrect category.
- The mapping is built from unique values, avoiding repeated fuzzy matching for identical values.

This makes the cleaning process both automated and auditable.

---

# 6. `DATA_LOAD.py`

## Purpose

Provide a reusable Excel ingestion function.

```python
def load_excel(file_path: str):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return pd.DataFrame()
```

The function keeps Excel loading consistent across the pipeline and provides basic failure handling.

The notebook therefore focuses on **pipeline orchestration** instead of repeating file-loading logic.

---

# 7. `DATA_TESTS.py`

This is the main validation module. It contains reusable checks for missing data, dates, business rules, and cross-dataset consistency.

---

## 7.1 `check_missing()`

```python
def check_missing(df, col_name):
```

The function checks whether a required column contains null values and returns a structured result containing:

- column name
- status
- message
- failed row indexes

The structured return format is important because the results can later be combined into a reporting table rather than only printed to the notebook.

Example:

```python
test1 = check_missing(df_1_DATA, "LA_LOB")
test2 = check_missing(df_1_DATA, "LA_LOSS_DATE")
```

---

## 7.2 `AS_AT_DATE_CLEAN()`

### Business Rule

Missing `LA_AS_AT_DATE` values are reconstructed from the corresponding payment date.

```text
Missing As-At Date
        ↓
Payment Date
        ↓
Payment Quarter
        ↓
Quarter-End
        ↓
Replacement As-At Date
```

The key implementation is:

```python
df_copy[as_at_col] = pd.to_datetime(
    df_copy[as_at_col],
    errors="coerce"
)

df_copy[payment_col] = pd.to_datetime(
    df_copy[payment_col],
    errors="coerce"
)

missing_mask = df_copy[as_at_col].isnull()

quarter_end_dates = (
    df_copy.loc[missing_mask, payment_col]
    .dt.to_period("Q")
    .dt.end_time
)

df_copy.loc[missing_mask, as_at_col] = quarter_end_dates
```

The function then runs the missing-value test again.

This is **rule-based imputation**, not generic filling: the replacement date is derived from the business relationship between payment date and reporting quarter.

The executed pipeline reported:

```text
4 missing As-At dates → substituted using payment-date quarter ends
```

Some issues remained because the function also preserves and reports unresolved cases.

---

## 7.3 `PAYMENT_DATE_CLEAN()`

### Business Rule

```text
Status = PAID
    → payment date should exist

Status ≠ PAID
    → payment date should not exist
```

The core logic:

```python
not_paid_mask = df_copy[status_col] != condition_value
paid_mask = df_copy[status_col] == condition_value

wiped_count = int(
    (not_paid_mask & df_copy[payment_date_col].notnull()).sum()
)

df_copy.loc[not_paid_mask, payment_date_col] = np.nan

missing_paid_dates = int(
    df_copy.loc[paid_mask, payment_date_col].isnull().sum()
)
```

The function separates two situations:

1. A payment date exists where it should not → clean it.
2. A `PAID` claim has no payment date → report it as unresolved.

Executed result:

```text
1 incorrect payment date removed
2 PAID claims still missing payment dates
```

Therefore the function can both **correct deterministic issues** and **surface issues requiring investigation**.

---

## 7.4 `PAYMENT_VS_LOSS_DATE_VALIDATE()`

### Business Rule

```text
Payment Date >= Loss Date
```

Same-day payment is allowed.

The validation first converts both fields to datetime and checks only rows where both dates exist:

```python
both_present_mask = (
    df_copy[payment_date_col].notnull()
    & df_copy[loss_date_col].notnull()
)

if allow_same_day:
    invalid_mask = (
        both_present_mask
        & (
            df_copy[payment_date_col]
            < df_copy[loss_date_col]
        )
    )
```

Instead of treating every missing date as an invalid record, the function distinguishes:

```text
VALID
INVALID_PAYMENT_BEFORE_LOSS
MISSING_DATE
```

It also writes the result back to the dataset:

```python
df_copy["PAYMENT_VS_LOSS_VALIDATION"] = np.where(
    ~both_present_mask,
    "MISSING_DATE",
    np.where(
        invalid_mask,
        "INVALID_PAYMENT_BEFORE_LOSS",
        "VALID"
    )
)
```

### Executed Result

```text
36 rows checked
47 rows skipped because of missing dates
0 invalid payment-before-loss records
```

This makes the validation **row-level and auditable**, not just a single PASS/FAIL statement.

---

## 7.5 `VALUATION_VS_LOSS_DATE_VALIDATE()`

This function extends the validation framework to a second dataset containing valuation dates.

### Workflow

```text
Valuation Dataset
      ↓
Check Key Duplicates
      ↓
Merge on Business Key
      ↓
Datetime Conversion
      ↓
Compare Valuation vs Loss Date
      ↓
Report Invalid / Missing / Unmatched Rows
```

The merge is:

```python
df_copy = df1_copy.merge(
    df2_copy[[key_col, valuation_date_col]],
    on=key_col,
    how="left"
)
```

Before the merge, duplicate keys are checked because duplicated keys can multiply rows during a join.

The implemented rule is:

```text
Loss Date <= Valuation Date
```

The function reports:

- invalid records
- missing dates
- unmatched records
- duplicate keys

This function was implemented but not executed in the shown pipeline because the required valuation dataset was not supplied.

---

## 7.6 `tests_agg()`

The individual test dictionaries are collected into one reporting structure:

```python
def tests_agg(**t):
    list_of_tests = []

    for key, value in t.items():
        list_of_tests.append(value)

    return list_of_tests
```

The notebook then converts the results into a DataFrame and exports:

```python
tests_table.to_excel(
    r"...\Report_Exported\Tests_Report.xlsx"
)
```

This creates a centralized QA output instead of leaving validation results scattered across notebook cells.

---

# 8. Financial Analysis in `Data_Pipeline.ipynb`

After cleaning and validation, the pipeline independently calculates financial aggregates from the claims data.

## 8.1 Status-Based Financial Aggregation

The dataset is grouped by `LA_LOB`.

Four measures are calculated:

```text
Gross Outstanding
Ceded Outstanding
Gross Paid
Ceded Paid
```

The aggregation applies claim status as the business condition:

```text
OS
→ Outstanding amounts

PAID
→ Paid amounts
```

Conceptually:

```text
Claim-Level Data
      ↓
Group by LOB
      ↓
Filter by Status
      ↓
Sum Gross / Ceded
      ↓
LOB Financial Summary
```

The calculated summary is renamed to a common `LOB` field and missing aggregates are normalized to `0.0`.

---

# 9. Financial Reconciliation

The independently calculated summary is compared with `5_FINANCIALS.xlsx`.

Reference columns:

```text
LOB
SUM_OF_GROSS_OS
SUM_OF_CEDED_OS
SUM_OF_GROSS_PAID
SUM_OF_CEDED_PAID
```

For each measure:

```python
df_summary["SUM_OF_GROSS_OS_diff"] = (
    df_summary["SUM_OF_GROSS_OS"]
    - df_5_FINANCIALS["SUM_OF_GROSS_OS"]
)
```

The same calculation is applied to the other three measures.

### Reconciliation Fields

```text
SUM_OF_GROSS_OS_diff
SUM_OF_CEDED_OS_diff
SUM_OF_GROSS_PAID_diff
SUM_OF_CEDED_PAID_diff
```

The interpretation is direct:

```text
Calculated Value - Reference Value = Difference
```

A zero difference indicates that the independently calculated aggregate matches the corresponding reference value.

The result is exported to:

```text
Financials_Test_Report.xlsx
```

This is a key part of the project because the financial totals are **recalculated from the underlying claims data rather than simply trusted from the existing report**.

---

# 10. Reporting Layer

The pipeline produces two final Excel outputs:

```text
Report_Exported/
│
├── Tests_Report.xlsx
└── Financials_Test_Report.xlsx
```

### `Tests_Report.xlsx`

Contains the structured results of:

- missing-value checks
- date cleaning checks
- payment-date validation
- payment-vs-loss validation
- other executed validation rules

### `Financials_Test_Report.xlsx`

Contains:

- calculated LOB-level financial aggregates
- reference financial values
- calculated-vs-reference differences

---

# 11. Engineering Design

The implementation follows a modular separation of responsibilities:

| Module | Responsibility |
|---|---|
| `DATA_LOAD.py` | Excel ingestion |
| `COLUMN_NAMES_CLEAN.py` | Schema alignment |
| `LOB_AND_STATUS_CLEAN.py` | Categorical standardization |
| `DATA_TESTS.py` | Data-quality and business-rule validation |
| `Data_Pipeline.ipynb` | Pipeline orchestration and financial analysis |
| `Report_Exported/` | Automated QA and reconciliation outputs |

The strongest engineering aspects are:

- **Reference-driven schema standardization**
- **Fuzzy matching with RapidFuzz**
- **Rule-based categorical cleaning**
- **Business-rule-driven date cleaning**
- **Reusable validation functions**
- **Row-level validation flags**
- **Cross-dataset validation**
- **Independent financial aggregation**
- **Financial reconciliation**
- **Modular Python architecture**
- **Automated Excel reporting**

---

# 12. End-to-End Logic

```text
                 RAW EXCEL SOURCES
                        │
                        ▼
                 DATA_LOAD.py
                        │
                        ▼
              COLUMN STANDARDIZATION
              ┌─────────┴─────────┐
              │                   │
        Synonym Mapping      RapidFuzz
              └─────────┬─────────┘
                        ▼
               STANDARDIZED SCHEMA
                        │
                        ▼
              LOB / STATUS CLEANING
                        │
                  RapidFuzz +
                Reference Values
                        │
                        ▼
                  DATE CLEANING
                        │
                        ▼
              DATA QUALITY TESTS
                        │
          ┌─────────────┼─────────────┐
          ▼             ▼             ▼
     Missing Data   Date Rules   Cross-Dataset
                                  Validation
          └─────────────┬─────────────┘
                        ▼
                VALIDATED DATASET
                        │
                        ▼
             LOB FINANCIAL AGGREGATION
                        │
                        ▼
              REFERENCE RECONCILIATION
                        │
              ┌─────────┴─────────┐
              ▼                   ▼
       Tests_Report.xlsx   Financials_Test_Report.xlsx
```

## Final Result

The project implements an **ETL-style insurance data processing and validation pipeline**:

**Ingest → Standardize → Clean → Validate → Aggregate → Reconcile → Report**

The emphasis is not only on manipulating Excel data, but on making the transformation process **repeatable, rule-driven, auditable, and modular**.
