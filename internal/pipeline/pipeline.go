// Package pipeline implements the pure transformation stages over
// (headers, data), plus the JSON writer. Stage order is fixed.
package pipeline

import (
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"sort"

	"github.com/zhide915/excel-to-json/internal/apperr"
	"github.com/zhide915/excel-to-json/internal/cell"
)

type RenamePair struct{ Old, New string }

// ExtractHeaders makes row headerRow (1-based) the column names, drops rows
// 1..headerRow, and pads/truncates data rows to the widest data row. Unnamed
// columns become col_K; duplicate names are an error.
func ExtractHeaders(g cell.Grid, headerRow int) ([]string, cell.Grid, error) {
	if headerRow < 1 {
		return nil, nil, apperr.Errf(apperr.HeaderRowBelowOne, "--header-row must be >= 1, got %d", headerRow)
	}
	if headerRow > len(g) {
		return nil, nil, apperr.Errf(apperr.HeaderRowOutOfRange,
			"--header-row %d exceeds the worksheet's %d row(s)", headerRow, len(g))
	}

	headerValues := g[headerRow-1]
	dataRows := g[headerRow:]

	// Target width: widest data row, or header width if no data rows.
	width := len(headerValues)
	if len(dataRows) > 0 {
		width = 0
		for _, row := range dataRows {
			if len(row) > width {
				width = len(row)
			}
		}
	}

	headers := make([]string, 0, width)
	for i := 0; i < width; i++ {
		var v any
		if i < len(headerValues) {
			v = headerValues[i]
		}
		headers = append(headers, headerName(v, i))
	}

	if dupes := Duplicates(headers); len(dupes) > 0 {
		return nil, nil, apperr.Errf(apperr.DuplicateHeaders, "duplicate header names: %s", apperr.QuoteList(dupes))
	}

	padded := make(cell.Grid, 0, len(dataRows))
	for _, row := range dataRows {
		nr := make([]any, width)
		copy(nr, row)
		padded = append(padded, nr)
	}
	return headers, padded, nil
}

func headerName(v any, i int) string {
	switch t := v.(type) {
	case nil:
		return fmt.Sprintf("col_%d", i)
	case string:
		if t == "" {
			return fmt.Sprintf("col_%d", i)
		}
		return t
	case bool:
		if t {
			return "true"
		}
		return "false"
	case float64:
		return string(cell.AppendJSONFloat(nil, t))
	default:
		return fmt.Sprint(t)
	}
}

// Duplicates returns the sorted, deduplicated set of names appearing more than
// once.
func Duplicates(items []string) []string {
	seen := make(map[string]bool, len(items))
	var dupes []string
	for _, s := range items {
		if seen[s] {
			dupes = append(dupes, s)
		}
		seen[s] = true
	}
	sort.Strings(dupes)
	return slices.Compact(dupes)
}

// FilterColumns restricts to the selected names in the given order; a nil
// selection passes through. Missing names are an error.
func FilterColumns(headers []string, data cell.Grid, selected []string) ([]string, cell.Grid, error) {
	if selected == nil {
		return headers, data, nil
	}

	var missing []string
	for _, name := range selected {
		if !slices.Contains(headers, name) {
			missing = append(missing, name)
		}
	}
	if len(missing) > 0 {
		return nil, nil, apperr.Errf(apperr.ColumnsMissing, "--columns references missing header(s): %s", apperr.QuoteList(missing))
	}

	indices := make([]int, len(selected))
	for i, name := range selected {
		indices[i] = slices.Index(headers, name)
	}

	newHeaders := make([]string, len(indices))
	for i, idx := range indices {
		newHeaders[i] = headers[idx]
	}
	newData := make(cell.Grid, 0, len(data))
	for _, row := range data {
		nr := make([]any, len(indices))
		for i, idx := range indices {
			nr[i] = row[idx]
		}
		newData = append(newData, nr)
	}
	return newHeaders, newData, nil
}

// RenameColumns applies old→new renames; a nil/empty mapping passes through.
// Missing sources and resulting name collisions are errors.
func RenameColumns(headers []string, data cell.Grid, mapping []RenamePair) ([]string, cell.Grid, error) {
	if len(mapping) == 0 {
		return headers, data, nil
	}

	var missing []string
	for _, p := range mapping {
		if !slices.Contains(headers, p.Old) {
			missing = append(missing, p.Old)
		}
	}
	if len(missing) > 0 {
		return nil, nil, apperr.Errf(apperr.RenameMissing, "--rename references missing column(s): %s", apperr.QuoteList(missing))
	}

	m := make(map[string]string, len(mapping))
	for _, p := range mapping {
		m[p.Old] = p.New
	}
	newHeaders := make([]string, len(headers))
	for i, h := range headers {
		if n, ok := m[h]; ok {
			newHeaders[i] = n
		} else {
			newHeaders[i] = h
		}
	}

	if collisions := Duplicates(newHeaders); len(collisions) > 0 {
		return nil, nil, apperr.Errf(apperr.RenameCollision, "--rename produces column-name collisions: %s", apperr.QuoteList(collisions))
	}
	return newHeaders, data, nil
}

// DropEmpty removes all-null rows first, then all-null columns.
func DropEmpty(headers []string, data cell.Grid, enabled bool) ([]string, cell.Grid) {
	if !enabled {
		return headers, data
	}

	filtered := make(cell.Grid, 0, len(data))
	for _, row := range data {
		if !cell.AllNull(row) {
			filtered = append(filtered, row)
		}
	}
	if len(filtered) == 0 {
		return headers, filtered
	}

	var keep []int
	for i := range headers {
		colAllNull := true
		for _, row := range filtered {
			if row[i] != nil {
				colAllNull = false
				break
			}
		}
		if !colAllNull {
			keep = append(keep, i)
		}
	}

	newHeaders := make([]string, len(keep))
	for i, idx := range keep {
		newHeaders[i] = headers[idx]
	}
	newData := make(cell.Grid, 0, len(filtered))
	for _, row := range filtered {
		nr := make([]any, len(keep))
		for i, idx := range keep {
			nr[i] = row[idx]
		}
		newData = append(newData, nr)
	}
	return newHeaders, newData
}

// EncodeRecords serializes to a compact JSON array of objects, keys in header
// order.
func EncodeRecords(headers []string, data cell.Grid) []byte {
	buf := []byte{'['}
	for i, row := range data {
		if i > 0 {
			buf = append(buf, ',')
		}
		buf = append(buf, '{')
		n := min(len(headers), len(row))
		for j := 0; j < n; j++ {
			if j > 0 {
				buf = append(buf, ',')
			}
			buf = cell.AppendJSONString(buf, headers[j])
			buf = append(buf, ':')
			buf = cell.AppendJSONValue(buf, row[j])
		}
		buf = append(buf, '}')
	}
	return append(buf, ']')
}

// WriteJSON writes the records to path as UTF-8 without BOM, creating parent
// directories and overwriting silently.
func WriteJSON(headers []string, data cell.Grid, path string) error {
	if parent := filepath.Dir(path); parent != "" && parent != "." {
		if err := os.MkdirAll(parent, 0o755); err != nil {
			return apperr.Errf(apperr.Write, "failed to write %s: %s", path, err)
		}
	}
	if err := os.WriteFile(path, EncodeRecords(headers, data), 0o644); err != nil {
		return apperr.Errf(apperr.Write, "failed to write %s: %s", path, err)
	}
	return nil
}
