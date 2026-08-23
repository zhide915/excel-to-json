# Excel to JSON Converter: Specification

**Version:** 1.0.0
**Last updated:** 2026-04-17

---

## 1. Overview

`excel-to-json` is a deterministic command-line tool that converts Excel spreadsheets (`.xlsx`, `.xls`) to compact JSON.

The tool performs a small, fixed pipeline: read the first worksheet, extract a header row, optionally filter and rename columns, optionally strip empty rows and columns, then write the result as a compact JSON array of records. It ships as a single self-contained binary with no runtime dependencies.

### 1.1 Design principles

- **Deterministic output.** Given the same input file and flags, the tool produces byte-identical output every time. This is load-bearing: automated consumers can diff results to detect drift.
- **Fail loud on misconfiguration.** A missing column or unknown rename key produces an error with a clear message, never a silent coercion. Silent fallback in a data-conversion tool is how incorrect data reaches downstream systems undetected.
- **Narrow by intent.** The tool transforms structure (headers, columns, empties), not content. It does not author formulas, compute aggregates, or handle multiple worksheets. Scope creep toward a general-purpose spreadsheet tool is explicitly out of scope.
- **YAGNI.** Features are added when a concrete use case demands them, not speculatively.

---

## 2. Command-Line Interface

### 2.1 Invocation

```
excel-to-json INPUT.xlsx OUTPUT.json [options]
```

| ID | Requirement |
|----|-------------|
| FR-CLI-01 | The command MUST be named `excel-to-json`. |
| FR-CLI-02 | The first positional argument MUST be the input file path (required). The file MUST exist. |
| FR-CLI-03 | The second positional argument MUST be the output file path (required). |
| FR-CLI-04 | `--help` MUST display usage with flag descriptions and one realistic example. |

### 2.2 Options

| ID | Flag | Meaning | Default |
|----|------|---------|---------|
| FR-CLI-05 | `--header-row N` | 1-based row number containing column headers | `1` |
| FR-CLI-06 | `--columns "a,b,c"` | Keep only these columns (by header name, in this order) | keep all |
| FR-CLI-07 | `--rename "old:new,old:new"` | Rename columns (applied after `--columns`), whitespace trimmed | none |
| FR-CLI-08 | `--ignore-empty` | Drop rows and columns that are entirely null | off |
| FR-CLI-09 | `--verbose` | Print pipeline narrative to stderr (for debugging) | off |
| FR-CLI-10 | `--help` | Print usage and exit 0 | none |

No other flags. Notably absent: `--sheet-name`, `--skip-rows`, `--range-spec`, `--headers`, `--preview`, `--pretty`, `--encoding`, `--use-config`, `--save-config`.

### 2.3 Delimiter rules

| ID | Requirement |
|----|-------------|
| FR-CLI-11 | In `--columns` and `--rename`, whitespace around names MUST be trimmed. |
| FR-CLI-12 | In `--columns`, header names MUST NOT contain commas (reserved as delimiter). |
| FR-CLI-13 | In `--rename`, names MUST NOT contain commas or colons (reserved as delimiters). |

---

## 3. Excel File Reading

| ID | Requirement |
|----|-------------|
| FR-READ-01 | The application MUST accept `.xlsx` files. |
| FR-READ-02 | The application MUST accept `.xls` files. |
| FR-READ-03 | The application MUST read the **first worksheet** in the workbook. No sheet selection is supported. |
| FR-READ-04 | The application MUST raise a clear error if the file does not exist, is not accessible, or contains no data. |
| FR-READ-05 | Merged cells MUST be handled by taking the value from the top-left cell of the merged area and duplicating it across all cells in that region. |
| FR-READ-06 | Cells containing formulas MUST be read as their computed values, not as formula text. |
| FR-READ-07 | Hidden rows and columns MUST be included in the output (treated identically to visible rows/columns). |
| FR-READ-08 | Files with unsupported extensions MUST be rejected with an error stating the extension and the accepted ones (`.xlsx`, `.xls`). |

---

## 4. Data Type Handling

| ID | Requirement |
|----|-------------|
| FR-TYPE-01 | String cell values MUST be serialized as JSON strings. |
| FR-TYPE-02 | Numeric cell values MUST be serialized as JSON numbers, preserving the precision provided by the spreadsheet engine (no rounding). |
| FR-TYPE-03 | Boolean cell values MUST be serialized as JSON `true` or `false`. |
| FR-TYPE-04 | Datetime cell values MUST be serialized as ISO 8601 strings (e.g., `"2026-04-16T14:30:00"`). Date-only values use `%Y-%m-%d` short form when the time component is midnight. |
| FR-TYPE-05 | Empty or missing cell values MUST be serialized as JSON `null`. |
| FR-TYPE-06 | Cell error values (`#REF!`, `#DIV/0!`, `#N/A`, `#VALUE!`, `#NAME?`, `#NULL!`, `#NUM!`, `#GETTING_DATA`, `#SPILL!`, `#CALC!`) MUST be serialized as JSON `null`. This applies both to the native error-variant and to plain-string cells that match these markers. |
| FR-TYPE-07 | Hyperlink cells MUST be serialized using their display text. |
| FR-TYPE-08 | Engine-specific numeric types (NaN, infinity, library-internal representations) MUST be converted to JSON-compatible values; NaN and infinity become `null`. |
| FR-TYPE-09 | Any other non-serializable type MUST fall back to its string representation. |

---

## 5. Processing Pipeline

Steps MUST be applied in exactly this order. Steps whose controlling flag is not set are skipped.

| Order | Step | Condition |
|-------|------|-----------|
| 1 | Read first worksheet (raw) | always |
| 2 | Extract headers at `--header-row`; drop all rows up to and including that row | always (default row = 1) |
| 3 | Filter columns per `--columns` | `--columns` set |
| 4 | Apply `--rename` | `--rename` set |
| 5 | Drop empty rows and columns | `--ignore-empty` set |
| 6 | Serialize and write JSON output | always |

### 5.1 Header extraction (step 2)

| ID | Requirement |
|----|-------------|
| FR-HDR-01 | `--header-row N` is a 1-based row number in the raw worksheet. That row's values become the column names. |
| FR-HDR-02 | All rows from row 1 through row N (inclusive) MUST be removed after header extraction. |
| FR-HDR-03 | If N exceeds the worksheet's row count, the application MUST produce a clear error stating the actual count. |
| FR-HDR-04 | If the header row has fewer non-null values than data columns, unnamed columns (trailing or middle) MUST be named `col_K` where K is the 0-based index of the unnamed column. |
| FR-HDR-05 | If the header row has more values than data columns, excess header values MUST be truncated. |
| FR-HDR-06 | Duplicate header names MUST be rejected with an error listing the duplicates. |

### 5.2 Column filtering (step 3)

| ID | Requirement |
|----|-------------|
| FR-COL-01 | When `--columns` is provided, output columns MUST be restricted to the listed names, in the order given. |
| FR-COL-02 | If any name in `--columns` is not present in the extracted headers, the application MUST reject with an error listing the missing name(s). No silent-ignore. |

### 5.3 Rename (step 4)

| ID | Requirement |
|----|-------------|
| FR-REN-01 | When `--rename` is provided, matching columns MUST be renamed from `old` to `new`. |
| FR-REN-02 | Rename is applied **after** column filtering: keys reference the post-filter column names. |
| FR-REN-03 | If any `old` key is not present in the current columns, the application MUST reject with an error listing the missing key(s). No silent-ignore. |
| FR-REN-04 | Columns not mentioned in `--rename` MUST retain their current name. |
| FR-REN-05 | If renaming would produce duplicate column names, the application MUST reject with a collision error. |

### 5.4 Empty removal (step 5)

| ID | Requirement |
|----|-------------|
| FR-EMPTY-01 | When `--ignore-empty` is set, rows where every remaining cell is `null` MUST be removed. |
| FR-EMPTY-02 | When `--ignore-empty` is set, columns where every remaining cell is `null` MUST be removed (after row removal). |

---

## 6. JSON Output

| ID | Requirement |
|----|-------------|
| FR-OUT-01 | The output MUST be a JSON array of objects (records format): each object is one row, keys are column names. |
| FR-OUT-02 | The output MUST be written as **compact** JSON (no extra whitespace). |
| FR-OUT-03 | The output MUST be written using **UTF-8** encoding. |
| FR-OUT-04 | Unicode characters MUST be preserved as-is (no `\uXXXX` escape sequences). |
| FR-OUT-05 | Output rows MUST appear in the same order as in the spreadsheet (after processing). |
| FR-OUT-06 | The output MUST be written to the `OUTPUT.json` path. Parent directories MUST be created automatically if missing. |
| FR-OUT-07 | If the output file already exists, it MUST be overwritten silently. |
| FR-OUT-08 | If all data is removed during processing (zero rows remain), the output MUST be an empty JSON array `[]`. |
| FR-OUT-09 | Stdout MUST NOT be used for the JSON payload or for success messages. Only `--verbose` (stderr) produces informational output. |
| FR-OUT-10 | Key order within each output object MUST match the final column order. This is required for determinism (NF-DET-01). |

---

## 7. Verbose Mode

| ID | Requirement |
|----|-------------|
| FR-VERB-01 | When `--verbose` is set, the application MUST print a step-by-step narrative of the pipeline to **stderr**. |
| FR-VERB-02 | Each pipeline step that executes MUST produce at least one verbose line naming the step and summarizing its effect. |
| FR-VERB-03 | Verbose output MUST be written to stderr only; stdout MUST remain unused. |
| FR-VERB-04 | When `--verbose` is NOT set, the application MUST produce no output on success (silent). Errors always go to stderr regardless of `--verbose`. |

### 7.1 Example verbose output

```
[1/5] Read sheet: 142 rows × 8 columns
[2/5] Header row 2: ["EmpID", "Name", "Dept", "Salary", "Bonus", "Notes", "Status", "Tax"]
      Data rows remaining: 140
[3/5] --columns filter: kept 2 columns ["EmpID", "Salary"]
[4/5] --rename: ["employee_id", "Salary"]
[5/5] --ignore-empty: dropped 3 empty rows, 0 empty columns. 137 records remain.
Wrote 137 records to out.json
```

Skipped stages are rendered as `[N/5] --flag: (skipped)`.

---

## 8. Validation

Pre-read validations (run before opening the file):

| ID | Rule | Error behavior |
|----|------|----------------|
| VR-01 | `INPUT.xlsx` must exist and be readable | Reject with file path |
| VR-02 | Input file extension must be `.xlsx` or `.xls` | Reject stating extension and accepted list |
| VR-03 | `OUTPUT.json` parent directory must be creatable (not an existing file) | Reject with parent path |
| VR-04 | `--header-row` must be >= 1 | Reject with explanation |
| VR-05 | `--columns` list must have no duplicates and no empty strings | Reject listing issues |
| VR-06 | `--rename` keys and values must be non-empty strings, no duplicate source keys | Reject with the offending pair |
| VR-07 | `--rename` must not produce duplicate destination names | Reject listing the collisions |

Post-read validations (depend on worksheet contents):

| ID | Rule | Error behavior |
|----|------|----------------|
| VR-08 | `--header-row` must not exceed the worksheet's row count | Reject with row count |
| VR-09 | Header row must not contain duplicate names | Reject listing duplicates |
| VR-10 | Every name in `--columns` must exist in the extracted headers | Reject listing missing names |
| VR-11 | Every `old` key in `--rename` must exist in the post-filter column names; no destination collisions | Reject listing missing keys or collisions |

---

## 9. Error Handling & Exit Codes

| ID | Requirement |
|----|-------------|
| NF-ERR-01 | Fail-fast: stop at the first error. |
| NF-ERR-02 | All errors MUST produce a clear, single-line message printed to stderr. |
| NF-ERR-03 | On any error, including CLI argument parse errors, the application MUST exit with exit code `1`. |
| NF-ERR-04 | On success, the application MUST exit with exit code `0`. `--help` and `--version` also exit `0`. |

### 9.1 Error scenarios

| Scenario | Expected behavior |
|----------|-------------------|
| Input file not found | Error with file path |
| Input file permission denied | Error with file path |
| Unsupported file extension | Error stating extension and accepted list |
| Empty file / no data after reading | Error stating no data found |
| Corrupt or invalid Excel file | Error with parse details |
| Output file permission denied | Error with output path |
| Output parent path is a file | Error stating the conflicting path |
| JSON serialization failure | Error with serialization details |

---

## 10. Non-Functional Requirements

| ID | Requirement |
|----|-------------|
| NF-DET-01 | Given the same input file and flags, the application MUST produce byte-identical output every time. |
| NF-COMPAT-01 | The application MUST run on Windows, macOS, and Linux. |
| NF-COMPAT-02 | The output MUST be valid JSON per RFC 8259. |
| NF-PERF-01 | The application MUST handle Excel files with tens of thousands of rows without crashing (memory-permitting). Verified via a 10k-row integration test (runs in ~0.8s on commodity hardware). |

---

## 11. Architecture

The implementation is organized into five small, single-purpose modules under `src/`:

| Module | Responsibility |
|--------|----------------|
| `main.rs` | Entry point. Parses args, calls `cli::run`, maps `Result` → `ExitCode`. ~20 lines. |
| `cli.rs` | clap-derive argument parsing, pre-read validation, and the `run()` orchestrator that chains pipeline stages. |
| `errors.rs` | Single `thiserror` enum covering every failure mode. Each variant's `Display` impl is the exact stderr message. |
| `types.rs` | Pure `cell_to_json(&calamine::Data) -> serde_json::Value` coercion. |
| `reader.rs` | Excel file reading. Dispatches by extension to `Xlsx` or `Xls` and expands merged cells. Both paths feed a shared `build_grid` helper. |
| `pipeline.rs` | Five pure functions (`extract_headers`, `filter_columns`, `rename_columns`, `drop_empty`, `to_records`) plus `write_json`. |

### 11.1 Data flow

```
INPUT.xlsx
    │
    ▼
  reader::read_excel  ──►  Grid = Vec<Vec<serde_json::Value>>
    │
    ▼
  pipeline::extract_headers  ──►  (Vec<String>, Vec<Vec<Value>>)
    │
    ▼
  pipeline::filter_columns   (if --columns)
    │
    ▼
  pipeline::rename_columns   (if --rename)
    │
    ▼
  pipeline::drop_empty       (if --ignore-empty)
    │
    ▼
  pipeline::to_records    ──►  Vec<serde_json::Map<String, Value>>
    │
    ▼
  pipeline::write_json    ──►  OUTPUT.json
```

Every stage except read/write is a pure function: no I/O, no global state. `cli::run` is the only place these are chained together.

### 11.2 Tech stack

- **Rust 1.75+** (edition 2021)
- **`calamine` 0.26** with `dates` feature: reads `.xlsx` and `.xls`
- **`clap` 4** (derive): argument parsing
- **`serde_json` 1** with `preserve_order` feature: JSON output with insertion-order keys (load-bearing for NF-DET-01 and FR-OUT-10)
- **`chrono` 0.4**: ISO 8601 date formatting
- **`thiserror` 1**: typed error enum

Dev-only: `assert_cmd`, `predicates`, `tempfile`, `rust_xlsxwriter`.

The release binary is ~1.4 MB on Windows (LTO + strip + codegen-units=1). Single-file self-contained distribution with no additional runtime on target machines.

---

## 12. Testing

70 tests across unit and integration layers:

| Layer | Count | Where |
|-------|-------|-------|
| CLI parsing & validation | 12 | `src/cli.rs` `#[cfg(test)]` |
| Cell coercion | 15 | `src/types.rs` `#[cfg(test)]` |
| Excel reading | 8 | `src/reader.rs` `#[cfg(test)]` |
| Pipeline stages + writer | 27 | `src/pipeline.rs` `#[cfg(test)]` |
| End-to-end via compiled binary | 8 + 1 ignored | `tests/integration.rs` |

The integration suite uses `assert_cmd` to invoke the compiled binary and covers: happy path, full pipeline with all flags combined, determinism (byte-identical repeat runs, NF-DET-01), unicode round-trip, missing-column error, quiet-by-default, verbose emission, 10k-row performance smoke (marked `#[ignore]`, runs with `cargo test -- --ignored` in ~0.8s, NF-PERF-01), and RFC 8259 JSON validity.

Excel fixtures are built dynamically at test time via `rust_xlsxwriter`, with no binary blobs committed to git.

### 12.1 Known testing gap

The `.xls` happy-path is **not** covered by automated tests. No pure-Rust `.xls` writer exists. The test `reader::tests::reads_xls_fixture_if_present` auto-skips when `tests/fixtures/simple.xls` is absent. Current `.xls` coverage:

- `corrupt_xls_errors`: exercises the parse-error branch (runs without fixture).
- Shared `build_grid` logic is format-agnostic and heavily tested via `.xlsx` fixtures.

**To close the gap:** create `tests/fixtures/simple.xls` with the content below, using Excel / LibreOffice Save As → Excel 97-2003 Workbook, or a Python + xlwt one-liner:

| ID | Name  |
|----|-------|
| 1  | Alice |
| 2  | Bob   |

The test runs automatically once the fixture is present.

---

## 13. Commands

```bash
cargo build              # Debug build
cargo build --release    # Optimized release binary (shipping artifact)
cargo test               # All 62 unit + 8 integration tests (~1s)
cargo test -- --ignored  # Plus 10k-row performance test (~1s)
cargo run -- INPUT.xlsx OUTPUT.json [options]  # Run the CLI via cargo
```

This is a **binary-only crate**: `cargo test --lib` errors with "no library targets found". Use `cargo test` without `--lib`.
