# Analysis File Builder

Upload **any CSV or Excel files**. The builder studies each one, finds the ID that links them, works out how each file should be merged, and produces **one verified analysis-ready table**. It exports **CSV, Excel, SPSS .sav**, plus an SPSS package that rebuilds the result in SPSS.

*Strategically powered by A1Intercept Technologies, Mysore · [aiintercept.online](https://aiintercept.online)*

## How it works
1. **Read everything**
   - **Formats:** CSV, TSV or TXT in any delimiter (comma, semicolon, tab, pipe) and encoding (UTF-8, Windows), plus `.xlsx` / `.xlsm`. Every non-empty sheet becomes a table, and a `.zip` of a folder works too.
   - **Header row:** found automatically, even with title lines above it.
   - **Clean-up:** formatted-but-empty rows and columns are dropped, unlabelled columns get a name, and duplicate headers are numbered.
2. **Study each table**
   - **Column types:** number, date (including text dates in many formats), date-time, time or text.
   - **Numbers stored as text** (e.g. `03`) are converted and counted.
   - **Ambiguous dates:** where both dd/mm and mm/dd fit, the column is flagged and day-first is assumed.
3. **Find the key**
   - **Scoring:** every column is scored as a possible ID on uniqueness, completeness, its name (`id`, `code`, `record`, `participant`, `subject`, `no`, …) and its value pattern.
   - **Matching:** IDs written differently still match: `001001` = `1001`, and ` p-01 ` = `P-01`.
   - **Base table:** the table whose unique, complete ID covers the most other tables. "master", "participant" or "registration" in the file name breaks ties, and copies are never chosen.
   - **Linking:** every other table is linked through whichever of its columns best matches the base IDs, even if that column has a different name. Tables matching fewer than 50% of IDs are reported as "not linked".
4. **Decide the grain**
   - **One row per ID:** merged directly.
   - **Several rows per ID:** the builder suggests a rule, which you can change:
     - **one column per visit/round:** when a column such as `visit` has a few levels and makes rows unique
     - **latest by date:** when the table has a date column
     - **first row**
     - **mean of numbers**
     - **keep as a separate long table:** exported on its own sheet and CSV
5. **Merge safely, with a log of every decision**
   - **Duplicate files:** byte-identical copies, or tables with the same IDs and values, are skipped.
   - **Shared columns:** a column appearing in several files with the same values is kept once, and missing values are filled from the other file.
   - **Conflicting columns:** if the same column name has different values, both are kept, with the file name added to the second.
   - **Copies of base columns:** text columns that copy a base-table column (e.g. a participant name typed again) are dropped if they agree 95% or more. Every mismatch goes to the **Consistency check**. Columns agreeing 70–95% are kept and labelled, with mismatches also reported.
   - **Orphan rows:** IDs not in the base table, rows with no ID, and exact duplicate rows are counted and removed.
   - **Added columns:**
     - `has_<table>` flags and `n_datasets`
     - `<id>_padded`, when IDs were written with leading zeros
6. **Reports**
   - **File profile:** for every table, the header row, empty rows dropped, link column, match %, grain, rule applied, text-numbers converted and date formats.
   - **Merge log, Consistency check and Outliers:** outliers are values with a robust z-score above 7.
   - **Not merged and Data dictionary.**
7. **Output options**
   - choose tables
   - de-identify (removes participant / family names, phone, address; staff names are kept)
   - keep only IDs present in every selected table
   - **HeLTI analysis variables:** the 73 verified HeLTI / EINSTEIN derived variables. This switches on automatically when HeLTI columns are present.

### Verification done during development
- **HeLTI baseline (32 files):** the builder chose the master list and `record_id` on its own and skipped all 7 copies. All 172 overlapping columns are identical to the verified hand-built merge (661,155 cells). The HeLTI variables match the verified analysis file. The one extra name mismatch is a phone number typed into a husband-name field.
- **Messy synthetic set:** title rows, a semicolon Windows-encoded CSV, `p-0001` vs `P-0001`, dd.mm.yyyy text dates, ambiguous dates, 3 visits per ID, repeat lab tests, an empty sheet, a byte-identical copy, a conflicting column and an unlinkable lookup table. Every rule was checked against the raw data.
- **SPSS package:** runs with 0 errors and every verification check PASS, and reproduces the builder's output value for value.

## Run on your own computer (most private: data never leaves the PC)
```
pip install -r requirements.txt
streamlit run app.py
```

## Host for free on Streamlit Community Cloud
1. Create a **private** GitHub repository and upload this folder (keep the structure). The `.gitignore` blocks data files.
2. Go to https://share.streamlit.io, sign in with GitHub, then choose **Create app**:
   - repository: your repo, main file `app.py`
   - **Advanced settings**: Python **3.12**
   - **Secrets**: add `APP_PASSWORD = "choose-a-strong-password"`
3. After it deploys, open **Settings → Sharing**, make the app private and invite your team.

Alternative: Hugging Face Spaces (Streamlit SDK, Private Space, `APP_PASSWORD` secret).

## Data protection
Research files often contain names and health data. Before cloud hosting, confirm that it is allowed by the study's ethics approval and data-sharing agreement. Always set `APP_PASSWORD`, and use **De-identify** or a local run when names aren't needed. Uploads live only in a temporary folder for the session, and **Clear session** deletes them.

## Branding
All branding lives in `branding.py`. Put your logo at `assets/logo.png` (or `.svg`); it replaces the placeholder. If you change the colours, also set `primaryColor` in `.streamlit/config.toml`.

## Code map
| Path | Purpose |
|---|---|
| `app.py` | Streamlit interface (upload → study → plan → build → download) |
| `pipeline/generic.py` | Reading, typing, key detection, planning, merging, reports, staging CSVs, SPSS syntax |
| `pipeline/generic_exports.py` | Output options, CSV / Excel / SPSS / package writers, HeLTI preset |
| `pipeline/derive.py`, `analysis_syntax.py`, `derived_labels.json` | HeLTI derived variables (Python + SPSS) |
| `branding.py`, `assets/`, `.streamlit/config.toml` | Branding and theme |
