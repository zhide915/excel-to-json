# CLAUDE.md: excel-to-json

Project-specific context that isn't obvious from the code. Read it before making non-trivial changes.

## What this is

A command-line converter. Reads an Excel file (`.xlsx`), applies a small pipeline (header row extraction, column filtering, rename, empty-row/column removal) and writes compact JSON. Deterministic output, single-binary deployment, no runtime dependencies.

`docs/SPECIFICATION.md` is the contract. **Do not modify it**.

## What this is not

- Not an interactive editor. No TUI, no prompts, no progress bar.
- Not a general-purpose spreadsheet tool. No formula authoring, no multi-sheet handling, no pivot tables.
- Not a streaming tool. The entire grid is materialized in memory (fine for tens of thousands of rows; not designed for millions).

## Tech stack

`go build` produces one self-contained executable. `github.com/xuri/excelize/v2` is the only external dependency. It reads `.xlsx` and generates the test fixtures. CLI parsing is stdlib `flag`; JSON output is a small hand-rolled encoder (see below).

## Known gap: `.xls` is not supported

FR-READ-02 requires `.xls`, but no maintained Go `.xls` reader exists. The CLI accepts the extension, but reading one errors, telling the user to convert to `.xlsx`. `TestXlsDeferredError` pins this. If a solid `.xls` library appears, wire it into the `readExcel` dispatch in `internal/reader`.

## Architecture

Thin `main.go` at the root delegating to `internal/` packages, one per responsibility; tests live alongside each package:

| Package | Role |
|---------|------|
| `main.go` | `os.Exit(cli.Main(os.Args[1:]))`. Nothing else. |
| `internal/cli` | Flag parsing, pre-read validation, the `run()` orchestrator with `--verbose` narration, exit-code mapping. |
| `internal/apperr` | `Error` with a `Kind` discriminator; each message is the exact stderr text. |
| `internal/cell` | The cell-value domain: `Grid` type, error markers, date detection, JSON scalar encoders. Pure, no I/O. |
| `internal/reader` | `.xlsx` reading via excelize: raw-value grid, used-range cropping, merged-cell expansion, per-cell type coercion. |
| `internal/pipeline` | Pure stages (`ExtractHeaders`, `FilterColumns`, `RenameColumns`, `DropEmpty`) plus `EncodeRecords`/`WriteJSON`. |
| `main_test.go` | Root-level end-to-end suite: builds the real binary in `TestMain`, execs it. |

The pipeline runs in a fixed order. See [docs/SPECIFICATION.md §5](docs/SPECIFICATION.md#5-processing-pipeline). Stages out of order produce different results.

Cell values are `any` restricted to `nil | bool | float64 | string`. There is no integer type: Excel numbers are all float64, and the encoder prints integral floats without a decimal point (`5000`).

## excelize quirks (why reader.go looks the way it does)

1. **`GetRows`/`GetCellValue` return *formatted* strings by default**: numbers per their numfmt, dates as locale strings. The reader always passes `excelize.Options{RawCellValue: true}` and re-types cells itself via `GetCellType`.
2. **There is no exported "is this cell a date?" API.** Date cells are just serials + a date number format. `cellHasDateStyle` resolves the cell's style and checks builtin date IDs (14–22, 27–36, 45–47, 50–58) or scans `CustomNumFmt` for y/m/d/h/s tokens (`cell.CustomFmtIsDateTime`). Decisions are cached per style index.
3. **`GetCellType` maps `t="str"` (string formula results) to `CellTypeFormula`**. Treat it as a string.
4. **The module path is `github.com/xuri/excelize/v2`**, not `qax-os/excelize` (that's the GitHub org alias; `go get` on it fails).
5. **encoding/json is not used for output**. It can't do ordered map keys and it HTML-escapes `<>&`. `pipeline.EncodeRecords` hand-writes compact JSON: header-order keys, minimal string escaping, and float formatting that replicates encoding/json's canonical algorithm.

## Running tests

```bash
go test ./...          # All tests, including the 10k-row performance test
go test -short ./...   # Skips the slow 10k-row test
go build               # Produces the shipping artifact
go vet ./... && gofmt -l .   # Keep clean
```

Integration tests build the real binary once in `TestMain` and exec it, with no test-only code paths. Excel fixtures are generated at test time with excelize; no binary blobs in git.

## Things to avoid

- **Don't silently bump excelize.** Its raw-value/cell-type/style behaviors are load-bearing for the reader; upgrade only as an explicit task and re-check the quirks above.
- **Resist scope creep.** Preview mode, config files, pretty-printing, and `--encoding` are out of scope by design. One shape, one output format, a handful of transformations.
- **Don't soften hard errors.** Missing rename keys and missing column names MUST error (FR-REN-03, VR-10). Silent coercion in a data-conversion tool lets bad data into downstream systems unnoticed.
- **Don't change the output format without updating tests.** `TestDeterminismByteIdentical`, `TestHappyPath`, and `TestOutputIsValidJSON` assert exact bytes and RFC 8259 round-trip.
- **Don't reword error messages casually.** Several tests pin the exact strings.
