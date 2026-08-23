package pipeline

import (
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/zhide915/excel-to-json/internal/apperr"
	"github.com/zhide915/excel-to-json/internal/cell"
)

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

// --- ExtractHeaders ---

func TestRowOneDefault(t *testing.T) {
	g := cell.Grid{
		{"a", "b", "c"},
		{1.0, 2.0, 3.0},
		{4.0, 5.0, 6.0},
	}
	h, d, err := ExtractHeaders(g, 1)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(h, []string{"a", "b", "c"}) {
		t.Fatalf("headers %v", h)
	}
	want := cell.Grid{{1.0, 2.0, 3.0}, {4.0, 5.0, 6.0}}
	if !reflect.DeepEqual(d, want) {
		t.Fatalf("data %v", d)
	}
}

func TestRowThreeDropsBanner(t *testing.T) {
	g := cell.Grid{
		{"Title row", nil, nil},
		{"subtitle", nil, nil},
		{"A", "B", "C"},
		{1.0, 2.0, 3.0},
	}
	h, d, err := ExtractHeaders(g, 3)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(h, []string{"A", "B", "C"}) {
		t.Fatalf("headers %v", h)
	}
	if !reflect.DeepEqual(d, cell.Grid{{1.0, 2.0, 3.0}}) {
		t.Fatalf("data %v", d)
	}
}

func TestHeaderRowOutOfRange(t *testing.T) {
	g := cell.Grid{{"a"}, {1.0}}
	_, _, err := ExtractHeaders(g, 5)
	ae := assertKind(t, err, apperr.HeaderRowOutOfRange)
	if ae.Error() != "--header-row 5 exceeds the worksheet's 2 row(s)" {
		t.Fatalf("message %q", ae.Error())
	}
}

func TestDuplicateHeadersError(t *testing.T) {
	g := cell.Grid{
		{"A", "B", "A"},
		{1.0, 2.0, 3.0},
	}
	_, _, err := ExtractHeaders(g, 1)
	ae := assertKind(t, err, apperr.DuplicateHeaders)
	if ae.Error() != `duplicate header names: ["A"]` {
		t.Fatalf("message %q", ae.Error())
	}
}

func TestShortHeaderPadsColNames(t *testing.T) {
	g := cell.Grid{
		{"A", "B", nil, nil},
		{1.0, 2.0, 3.0, 4.0},
	}
	h, _, err := ExtractHeaders(g, 1)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(h, []string{"A", "B", "col_2", "col_3"}) {
		t.Fatalf("headers %v", h)
	}
}

func TestMiddleNullHeaderBecomesColName(t *testing.T) {
	g := cell.Grid{
		{"A", nil, "C"},
		{1.0, 2.0, 3.0},
	}
	h, _, err := ExtractHeaders(g, 1)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(h, []string{"A", "col_1", "C"}) {
		t.Fatalf("headers %v", h)
	}
}

func TestNonStringHeadersCoercedToString(t *testing.T) {
	g := cell.Grid{
		{1.0, 2.5, true},
		{"a", "b", "c"},
	}
	h, _, err := ExtractHeaders(g, 1)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(h, []string{"1", "2.5", "true"}) {
		t.Fatalf("headers %v", h)
	}
}

// --- FilterColumns ---

func TestFilterNonePassesThrough(t *testing.T) {
	h := []string{"a", "b", "c"}
	d := cell.Grid{{1.0, 2.0, 3.0}}
	nh, nd, err := FilterColumns(h, d, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(nh, h) || !reflect.DeepEqual(nd, d) {
		t.Fatalf("changed: %v %v", nh, nd)
	}
}

func TestFilterKeepsSelectedInOrder(t *testing.T) {
	h := []string{"a", "b", "c"}
	d := cell.Grid{{1.0, 2.0, 3.0}}
	nh, nd, err := FilterColumns(h, d, []string{"c", "a"})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(nh, []string{"c", "a"}) {
		t.Fatalf("headers %v", nh)
	}
	if !reflect.DeepEqual(nd, cell.Grid{{3.0, 1.0}}) {
		t.Fatalf("data %v", nd)
	}
}

func TestFilterMissingColumnErrors(t *testing.T) {
	h := []string{"a", "b"}
	d := cell.Grid{{1.0, 2.0}}
	_, _, err := FilterColumns(h, d, []string{"a", "z"})
	ae := assertKind(t, err, apperr.ColumnsMissing)
	if ae.Error() != `--columns references missing header(s): ["z"]` {
		t.Fatalf("message %q", ae.Error())
	}
}

func TestFilterMultipleMissingListed(t *testing.T) {
	h := []string{"a"}
	d := cell.Grid{{1.0}}
	_, _, err := FilterColumns(h, d, []string{"x", "y"})
	ae := assertKind(t, err, apperr.ColumnsMissing)
	if ae.Error() != `--columns references missing header(s): ["x", "y"]` {
		t.Fatalf("message %q", ae.Error())
	}
}

func TestFilterEmptyData(t *testing.T) {
	h := []string{"a", "b"}
	nh, nd, err := FilterColumns(h, cell.Grid{}, []string{"b"})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(nh, []string{"b"}) || len(nd) != 0 {
		t.Fatalf("got %v %v", nh, nd)
	}
}

// --- RenameColumns ---

func TestRenameNonePassesThrough(t *testing.T) {
	h := []string{"a", "b"}
	d := cell.Grid{{1.0, 2.0}}
	nh, nd, err := RenameColumns(h, d, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(nh, h) || !reflect.DeepEqual(nd, d) {
		t.Fatalf("changed: %v %v", nh, nd)
	}
}

func TestRenameSingle(t *testing.T) {
	h := []string{"a", "b"}
	d := cell.Grid{{1.0, 2.0}}
	nh, _, err := RenameColumns(h, d, []RenamePair{{Old: "a", New: "alpha"}})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(nh, []string{"alpha", "b"}) {
		t.Fatalf("headers %v", nh)
	}
}

func TestRenameMissingSourceErrors(t *testing.T) {
	h := []string{"a", "b"}
	d := cell.Grid{{1.0, 2.0}}
	_, _, err := RenameColumns(h, d, []RenamePair{{Old: "missing", New: "x"}})
	ae := assertKind(t, err, apperr.RenameMissing)
	if ae.Error() != `--rename references missing column(s): ["missing"]` {
		t.Fatalf("message %q", ae.Error())
	}
}

func TestRenameCollisionErrors(t *testing.T) {
	h := []string{"a", "b"}
	d := cell.Grid{{1.0, 2.0}}
	_, _, err := RenameColumns(h, d, []RenamePair{{Old: "a", New: "same"}, {Old: "b", New: "same"}})
	assertKind(t, err, apperr.RenameCollision)
}

// --- DropEmpty ---

func TestDropEmptyDisabledPassesThrough(t *testing.T) {
	h := []string{"a", "b"}
	d := cell.Grid{{nil, nil}, {1.0, 2.0}}
	nh, nd := DropEmpty(h, d, false)
	if !reflect.DeepEqual(nh, h) || !reflect.DeepEqual(nd, d) {
		t.Fatalf("changed: %v %v", nh, nd)
	}
}

func TestDropEmptyRemovesAllNullRow(t *testing.T) {
	h := []string{"a", "b"}
	d := cell.Grid{
		{1.0, 2.0},
		{nil, nil},
		{3.0, 4.0},
	}
	nh, nd := DropEmpty(h, d, true)
	if !reflect.DeepEqual(nh, []string{"a", "b"}) {
		t.Fatalf("headers %v", nh)
	}
	if len(nd) != 2 {
		t.Fatalf("rows %v", nd)
	}
}

func TestDropEmptyRemovesAllNullColumn(t *testing.T) {
	h := []string{"a", "b", "c"}
	d := cell.Grid{
		{1.0, nil, 3.0},
		{4.0, nil, 6.0},
	}
	nh, nd := DropEmpty(h, d, true)
	if !reflect.DeepEqual(nh, []string{"a", "c"}) {
		t.Fatalf("headers %v", nh)
	}
	if !reflect.DeepEqual(nd, cell.Grid{{1.0, 3.0}, {4.0, 6.0}}) {
		t.Fatalf("data %v", nd)
	}
}

func TestDropEmptyRowsThenColumns(t *testing.T) {
	h := []string{"a", "b", "c"}
	d := cell.Grid{
		{1.0, nil, 3.0},
		{nil, nil, nil},
		{4.0, nil, 6.0},
	}
	nh, nd := DropEmpty(h, d, true)
	if !reflect.DeepEqual(nh, []string{"a", "c"}) {
		t.Fatalf("headers %v", nh)
	}
	if !reflect.DeepEqual(nd, cell.Grid{{1.0, 3.0}, {4.0, 6.0}}) {
		t.Fatalf("data %v", nd)
	}
}

func TestDropEmptyAllRowsRemoved(t *testing.T) {
	h := []string{"a", "b"}
	d := cell.Grid{{nil, nil}, {nil, nil}}
	nh, nd := DropEmpty(h, d, true)
	if !reflect.DeepEqual(nh, []string{"a", "b"}) || len(nd) != 0 {
		t.Fatalf("got %v %v", nh, nd)
	}
}

// --- EncodeRecords / WriteJSON ---

func TestRecordsPreserveKeyOrder(t *testing.T) {
	h := []string{"z", "a", "m"}
	d := cell.Grid{{1.0, 2.0, 3.0}}
	got := string(EncodeRecords(h, d))
	if got != `[{"z":1,"a":2,"m":3}]` {
		t.Fatalf("got %q", got)
	}
}

func TestWriteCompactJSON(t *testing.T) {
	out := filepath.Join(t.TempDir(), "out.json")
	if err := WriteJSON([]string{"a"}, cell.Grid{{1.0}}, out); err != nil {
		t.Fatal(err)
	}
	text, _ := os.ReadFile(out)
	if string(text) != `[{"a":1}]` {
		t.Fatalf("got %q", text)
	}
}

func TestWriteEmptyArray(t *testing.T) {
	out := filepath.Join(t.TempDir(), "out.json")
	if err := WriteJSON([]string{"a"}, cell.Grid{}, out); err != nil {
		t.Fatal(err)
	}
	text, _ := os.ReadFile(out)
	if string(text) != "[]" {
		t.Fatalf("got %q", text)
	}
}

func TestWriteCreatesParentDirs(t *testing.T) {
	out := filepath.Join(t.TempDir(), "a", "b", "c", "out.json")
	if err := WriteJSON([]string{"x"}, cell.Grid{{1.0}}, out); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(out); err != nil {
		t.Fatal(err)
	}
}

func TestWriteOverwritesExisting(t *testing.T) {
	out := filepath.Join(t.TempDir(), "out.json")
	if err := os.WriteFile(out, []byte("old content"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := WriteJSON([]string{"new"}, cell.Grid{{true}}, out); err != nil {
		t.Fatal(err)
	}
	text, _ := os.ReadFile(out)
	if string(text) != `[{"new":true}]` {
		t.Fatalf("got %q", text)
	}
}

func TestUnicodePreservedNotEscaped(t *testing.T) {
	out := filepath.Join(t.TempDir(), "out.json")
	if err := WriteJSON([]string{"name"}, cell.Grid{{"田中 한국 测试"}}, out); err != nil {
		t.Fatal(err)
	}
	text, _ := os.ReadFile(out)
	s := string(text)
	if !strings.Contains(s, "田中") || strings.Contains(s, `\u`) {
		t.Fatalf("got %q", s)
	}
}

func TestNoBOMWritten(t *testing.T) {
	out := filepath.Join(t.TempDir(), "out.json")
	if err := WriteJSON([]string{"a"}, cell.Grid{{1.0}}, out); err != nil {
		t.Fatal(err)
	}
	raw, _ := os.ReadFile(out)
	if len(raw) >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF {
		t.Fatal("BOM written")
	}
}
