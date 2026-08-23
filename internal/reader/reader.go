// Package reader reads Excel files into a rectangular grid of cell values.
// .xlsx is read via excelize; .xls is not supported (no maintained Go .xls
// reader exists) and produces a clear error.
package reader

import (
	"math"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"

	"github.com/zhide915/excel-to-json/internal/apperr"
	"github.com/zhide915/excel-to-json/internal/cell"
)

func ReadExcel(path string) (cell.Grid, error) {
	ext := strings.TrimPrefix(strings.ToLower(filepath.Ext(path)), ".")
	switch ext {
	case "xlsx":
		return readXlsx(path)
	case "xls":
		return nil, apperr.Errf(apperr.XlsUnsupported,
			".xls support is not available in this build; convert %s to .xlsx", path)
	default:
		return nil, apperr.Errf(apperr.UnsupportedExtension,
			"unsupported extension %s; accepted: .xlsx, .xls", strconv.Quote(ext))
	}
}

func readXlsx(path string) (cell.Grid, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, apperr.Errf(apperr.Parse, "failed to parse %s: %s", path, err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, apperr.Errf(apperr.NoData, "no data in %s", path)
	}
	sheet := sheets[0]

	// Raw values preserve types; formatted values would stringify numbers and
	// dates per number format.
	rawRows, err := f.GetRows(sheet, excelize.Options{RawCellValue: true})
	if err != nil {
		return nil, apperr.Errf(apperr.Parse, "failed to parse %s: %s", path, err)
	}

	// Crop to the bounding box of non-empty cells: leading empty rows and
	// columns are excluded from the output.
	minR, minC, maxR, maxC := -1, -1, -1, -1
	for r, row := range rawRows {
		for c, v := range row {
			if v == "" {
				continue
			}
			if minR == -1 {
				minR, minC, maxR, maxC = r, c, r, c
				continue
			}
			if c < minC {
				minC = c
			}
			if c > maxC {
				maxC = c
			}
			maxR = r
		}
	}
	if minR == -1 {
		return nil, apperr.Errf(apperr.NoData, "no data in %s", path)
	}

	props, propsErr := f.GetWorkbookProps()
	cr := &cellReader{
		f:          f,
		sheet:      sheet,
		date1904:   propsErr == nil && props.Date1904 != nil && *props.Date1904,
		dateStyles: map[int]bool{},
	}

	g := make(cell.Grid, 0, maxR-minR+1)
	for r := minR; r <= maxR; r++ {
		row := make([]any, 0, maxC-minC+1)
		for c := minC; c <= maxC; c++ {
			raw := ""
			if r < len(rawRows) && c < len(rawRows[r]) {
				raw = rawRows[r][c]
			}
			v, err := cr.cellValue(r, c, raw)
			if err != nil {
				return nil, apperr.Errf(apperr.Parse, "failed to parse %s: %s", path, err)
			}
			row = append(row, v)
		}
		g = append(g, row)
	}

	// Duplicate each merged region's top-left value across the whole region.
	// Merge coordinates are absolute; the grid is cropped.
	merges, err := f.GetMergeCells(sheet)
	if err != nil {
		return nil, apperr.Errf(apperr.Parse, "failed to parse %s: %s", path, err)
	}
	for _, m := range merges {
		sc, sr, err1 := excelize.CellNameToCoordinates(m.GetStartAxis())
		ec, er, err2 := excelize.CellNameToCoordinates(m.GetEndAxis())
		if err1 != nil || err2 != nil {
			continue
		}
		sr, sc, er, ec = sr-1, sc-1, er-1, ec-1 // 1-based -> 0-based absolute
		var topLeft any
		if sr >= minR && sr <= maxR && sc >= minC && sc <= maxC {
			topLeft = g[sr-minR][sc-minC]
		}
		for r := max(sr, minR); r <= min(er, maxR); r++ {
			for c := max(sc, minC); c <= min(ec, maxC); c++ {
				g[r-minR][c-minC] = topLeft
			}
		}
	}

	// A sheet whose only content coerced to null (e.g. all error cells) counts
	// as empty.
	allEmpty := true
	for _, row := range g {
		if !cell.AllNull(row) {
			allEmpty = false
			break
		}
	}
	if allEmpty {
		return nil, apperr.Errf(apperr.NoData, "no data in %s", path)
	}

	return g, nil
}

type cellReader struct {
	f          *excelize.File
	sheet      string
	date1904   bool
	dateStyles map[int]bool // style index -> renders as date/time
}

// cellValue coerces one cell (0-based coordinates, raw stored value).
func (cr *cellReader) cellValue(r, c int, raw string) (any, error) {
	if raw == "" {
		return nil, nil
	}
	cellName, err := excelize.CoordinatesToCellName(c+1, r+1)
	if err != nil {
		return nil, err
	}
	ct, err := cr.f.GetCellType(cr.sheet, cellName)
	if err != nil {
		return nil, err
	}
	switch ct {
	case excelize.CellTypeBool:
		return raw == "1" || strings.EqualFold(raw, "true"), nil
	case excelize.CellTypeError:
		return nil, nil
	case excelize.CellTypeSharedString, excelize.CellTypeInlineString,
		excelize.CellTypeFormula: // t="str": a formula's cached string result
		if cell.IsErrorString(raw) {
			return nil, nil
		}
		return raw, nil
	case excelize.CellTypeDate:
		return raw, nil // t="d" cells already store an ISO 8601 string
	default: // number cells, and numeric formula results (t="n" or absent)
		if cell.IsErrorString(raw) {
			return nil, nil
		}
		f64, perr := strconv.ParseFloat(raw, 64)
		if perr != nil {
			return raw, nil
		}
		if math.IsNaN(f64) || math.IsInf(f64, 0) {
			return nil, nil
		}
		isDate, err := cr.cellHasDateStyle(cellName)
		if err != nil {
			return nil, err
		}
		if isDate {
			t, terr := excelize.ExcelDateToTime(f64, cr.date1904)
			if terr != nil {
				return nil, nil
			}
			// Serial fractions carry float error; snap to the nearest second.
			return cell.FormatDateTime(t.Round(time.Second)), nil
		}
		return f64, nil
	}
}

func (cr *cellReader) cellHasDateStyle(cellName string) (bool, error) {
	styleID, err := cr.f.GetCellStyle(cr.sheet, cellName)
	if err != nil {
		return false, err
	}
	if v, ok := cr.dateStyles[styleID]; ok {
		return v, nil
	}
	isDate := false
	if style, serr := cr.f.GetStyle(styleID); serr == nil && style != nil {
		if style.CustomNumFmt != nil {
			isDate = cell.CustomFmtIsDateTime(*style.CustomNumFmt)
		} else {
			isDate = cell.IsBuiltinDateNumFmt(style.NumFmt)
		}
	}
	cr.dateStyles[styleID] = isDate
	return isDate, nil
}
