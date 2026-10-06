# excel-to-json

A deterministic command-line tool that converts Excel spreadsheets (`.xlsx`) to compact JSON.

Single self-contained binary with no runtime dependencies. UTF-8 preserved, keys in column order, byte-identical output on repeat runs, safe to chain into automated pipelines.

> **Note:** `.xls` (Excel 97-2003) is not supported: no maintained Go `.xls` reader exists. The CLI accepts the extension but errors, telling you to convert the file to `.xlsx`.

---

## Quick Start

```bash
excel-to-json INPUT.xlsx OUTPUT.json [options]
```

### Examples

```bash
# Minimum: first worksheet, row 1 = headers, all columns kept
excel-to-json report.xlsx out.json

# Skip banner rows; keep specific columns; rename them
excel-to-json data.xlsx out.json \
  --header-row 2 \
  --columns "ID,Name,Amount" \
  --rename "ID:record_id,Amount:value" \
  --ignore-empty

# Debug what each pipeline stage did (stderr)
excel-to-json report.xlsx out.json --verbose
```

---

## Flags

| Flag | Meaning | Default |
|------|---------|---------|
| `--header-row N` | 1-based row number containing headers | `1` |
| `--columns "a,b,c"` | Keep only these columns (by header name, in this order) | keep all |
| `--rename "old:new,old:new"` | Rename columns after filtering | none |
| `--ignore-empty` | Drop rows and columns that are entirely null | off |
| `--verbose` | Print step-by-step narrative to stderr | off |
| `--help` | Print usage | none |

See [docs/SPECIFICATION.md](docs/SPECIFICATION.md) for the full reference.

---

## Building from source

```bash
# Build the binary
go build

# Run the full test suite
go test ./...

# Skip the slow 10k-row performance test
go test -short ./...
```

Requires **Go 1.26+**.

---

## Output format

JSON array of objects, one object per row, keys in left-to-right column order:

```json
[{"ID":1,"Name":"Alice","Amount":5000},{"ID":2,"Name":"Bob","Amount":6000}]
```

Always compact (no whitespace), always UTF-8, Unicode preserved (no `\uXXXX` escaping). Written directly to the output path (parent directories created if missing; existing file overwritten silently).

On any error: single-line message to stderr, exit code `1`. On success: exit `0`, silent unless `--verbose`.

---

## Documentation

- [docs/SPECIFICATION.md](docs/SPECIFICATION.md): full requirements
