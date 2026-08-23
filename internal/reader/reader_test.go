package reader

import (
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"

	"github.com/zhide915/excel-to-json/internal/apperr"
	"github.com/zhide915/excel-to-json/internal/cell"
)

const sheet1 = "Sheet1"

func assertKind(t *testing.T, err error, kind apperr.Kind) *apperr.Error {
	t.Helper()
	var ae *apperr.Error
	if !errors.As(err, &ae) {
		t.Fatalf("expected *apperr.Error, got %T: %v", err, err)
	}
	if ae.Kind != kind {
		t.Fatalf("expected kind %d, got %d (%v)", kind, ae.Kind, ae)
	}
	return ae
}

func saveWorkbook(t *testing.T, f *excelize.File, dir, name string) string {
	t.Helper()
	path := filepath.Join(dir, name)
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}
	return path
}

func makeXlsx(t *testing.T, dir, name string, rows [][]string) string {
	t.Helper()
	f := excelize.NewFile()
	for r, row := range rows {
		for c, val := range row {
			cellName, _ := excelize.CoordinatesToCellName(c+1, r+1)
			if err := f.SetCellStr(sheet1, cellName, val); err != nil {
				t.Fatal(err)
			}
		}
	}
	return saveWorkbook(t, f, dir, name+".xlsx")
}

func makeXlsxMixed(t *testing.T, dir, name string) string {
	t.Helper()
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "ID")
	f.SetCellStr(sheet1, "B1", "Name")
	f.SetCellStr(sheet1, "C1", "Salary")
	f.SetCellValue(sheet1, "A2", 1)
	f.SetCellStr(sheet1, "B2", "Alice")
	f.SetCellValue(sheet1, "C2", 5000)
	f.SetCellValue(sheet1, "A3", 2)
	f.SetCellStr(sheet1, "B3", "Bob")
	f.SetCellValue(sheet1, "C3", 6000)
	return saveWorkbook(t, f, dir, name+".xlsx")
}

func TestReadsSimpleGrid(t *testing.T) {
	path := makeXlsxMixed(t, t.TempDir(), "simple")
	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	want := cell.Grid{
		{"ID", "Name", "Salary"},
		{1.0, "Alice", 5000.0},
		{2.0, "Bob", 6000.0},
	}
	if !reflect.DeepEqual(g, want) {
		t.Fatalf("grid %#v", g)
	}
}

func TestReaderUnicodePreserved(t *testing.T) {
	path := makeXlsx(t, t.TempDir(), "u", [][]string{{"名前"}, {"田中"}, {"한국"}})
	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	if g[1][0] != "田中" || g[2][0] != "한국" {
		t.Fatalf("grid %#v", g)
	}
}

func TestOnlyReadsFirstSheet(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "first_sheet")
	if _, err := f.NewSheet("Other"); err != nil {
		t.Fatal(err)
	}
	f.SetCellStr("Other", "A1", "second_sheet")
	path := saveWorkbook(t, f, t.TempDir(), "multi.xlsx")

	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(g, cell.Grid{{"first_sheet"}}) {
		t.Fatalf("grid %#v", g)
	}
}

func TestMergedCellsExpandTopLeft(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "H1")
	f.SetCellStr(sheet1, "B1", "H2")
	f.SetCellStr(sheet1, "C1", "H3")
	f.SetCellStr(sheet1, "A2", "region_a")
	if err := f.MergeCell(sheet1, "A2", "C2"); err != nil {
		t.Fatal(err)
	}
	f.SetCellValue(sheet1, "A3", 1)
	f.SetCellValue(sheet1, "B3", 2)
	f.SetCellValue(sheet1, "C3", 3)
	path := saveWorkbook(t, f, t.TempDir(), "merged.xlsx")

	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(g[1], []any{"region_a", "region_a", "region_a"}) {
		t.Fatalf("merged row %#v", g[1])
	}
}

func TestEmptyFileErrors(t *testing.T) {
	f := excelize.NewFile()
	path := saveWorkbook(t, f, t.TempDir(), "empty.xlsx")
	_, err := ReadExcel(path)
	assertKind(t, err, apperr.NoData)
}

func TestCorruptFileErrors(t *testing.T) {
	path := filepath.Join(t.TempDir(), "corrupt.xlsx")
	if err := os.WriteFile(path, []byte("not a real xlsx"), 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := ReadExcel(path)
	assertKind(t, err, apperr.Parse)
}

func TestXlsDeferredError(t *testing.T) {
	path := filepath.Join(t.TempDir(), "legacy.xls")
	if err := os.WriteFile(path, []byte("anything"), 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := ReadExcel(path)
	ae := assertKind(t, err, apperr.XlsUnsupported)
	if !strings.Contains(ae.Error(), ".xls support is not available in this build") ||
		!strings.Contains(ae.Error(), "convert") {
		t.Fatalf("message %q", ae.Error())
	}
}

func TestBoolCells(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "flag")
	f.SetCellBool(sheet1, "A2", true)
	f.SetCellBool(sheet1, "A3", false)
	path := saveWorkbook(t, f, t.TempDir(), "bool.xlsx")

	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	if g[1][0] != true {
		t.Fatalf("A2 = %#v", g[1][0])
	}
	// FALSE bool cells store raw "0"; must come back as bool, not number.
	if g[2][0] != false {
		t.Fatalf("A3 = %#v", g[2][0])
	}
}

func TestErrorMarkerStringsBecomeNull(t *testing.T) {
	path := makeXlsx(t, t.TempDir(), "err", [][]string{
		{"a", "b"},
		{"#N/A", "ok"},
	})
	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	if g[1][0] != nil || g[1][1] != "ok" {
		t.Fatalf("row %#v", g[1])
	}
}

func TestNumericLookingTextStaysString(t *testing.T) {
	path := makeXlsx(t, t.TempDir(), "txt", [][]string{
		{"code"},
		{"00123"},
	})
	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	if g[1][0] != "00123" {
		t.Fatalf("cell %#v", g[1][0])
	}
}

func TestDateTimeCells(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "when")
	f.SetCellValue(sheet1, "A2", time.Date(2026, 4, 16, 14, 30, 0, 0, time.UTC))
	f.SetCellValue(sheet1, "A3", time.Date(2026, 4, 16, 0, 0, 0, 0, time.UTC))
	path := saveWorkbook(t, f, t.TempDir(), "dates.xlsx")

	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	if g[1][0] != "2026-04-16T14:30:00" {
		t.Fatalf("A2 = %#v", g[1][0])
	}
	if g[2][0] != "2026-04-16" {
		t.Fatalf("A3 = %#v", g[2][0])
	}
}

func TestCustomDateFormatDetected(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "when")
	f.SetCellValue(sheet1, "A2", 45000) // serial for 2023-03-15
	custom := "yyyy-mm-dd"
	styleID, err := f.NewStyle(&excelize.Style{CustomNumFmt: &custom})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellStyle(sheet1, "A2", "A2", styleID); err != nil {
		t.Fatal(err)
	}
	// A number cell with a non-date custom format must stay numeric.
	f.SetCellValue(sheet1, "A3", 45000)
	money := "#,##0.00"
	moneyID, err := f.NewStyle(&excelize.Style{CustomNumFmt: &money})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellStyle(sheet1, "A3", "A3", moneyID); err != nil {
		t.Fatal(err)
	}
	path := saveWorkbook(t, f, t.TempDir(), "custom.xlsx")

	g, err := ReadExcel(path)
	if err != nil {
		t.Fatal(err)
	}
	if g[1][0] != "2023-03-15" {
		t.Fatalf("A2 = %#v", g[1][0])
	}
	if g[2][0] != 45000.0 {
		t.Fatalf("A3 = %#v", g[2][0])
	}
}
