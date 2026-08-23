package cli

import (
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/zhide915/excel-to-json/internal/apperr"
	"github.com/zhide915/excel-to-json/internal/pipeline"
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

func validInput(t *testing.T, dir string) string {
	t.Helper()
	p := filepath.Join(dir, "in.xlsx")
	if err := os.WriteFile(p, []byte("dummy"), 0o644); err != nil {
		t.Fatal(err)
	}
	return p
}

func TestParsesMinimumArgs(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	args, err := parseArgs([]string{inp, out})
	if err != nil {
		t.Fatal(err)
	}
	if args.headerRow != 1 || args.columns != nil || args.rename != nil ||
		args.ignoreEmpty || args.verbose {
		t.Fatalf("unexpected defaults: %+v", args)
	}
}

func TestParsesAllFlags(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	args, err := parseArgs([]string{
		inp, out,
		"--header-row", "2",
		"--columns", "A,B, C ",
		"--rename", "A:alpha, B : beta ",
		"--ignore-empty",
		"--verbose",
	})
	if err != nil {
		t.Fatal(err)
	}
	if args.headerRow != 2 {
		t.Fatalf("headerRow %d", args.headerRow)
	}
	if !reflect.DeepEqual(args.columns, []string{"A", "B", "C"}) {
		t.Fatalf("columns %v", args.columns)
	}
	want := []pipeline.RenamePair{{Old: "A", New: "alpha"}, {Old: "B", New: "beta"}}
	if !reflect.DeepEqual(args.rename, want) {
		t.Fatalf("rename %v", args.rename)
	}
	if !args.ignoreEmpty || !args.verbose {
		t.Fatal("bool flags not set")
	}
}

func TestRejectsMissingInput(t *testing.T) {
	dir := t.TempDir()
	inp := filepath.Join(dir, "nope.xlsx")
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out})
	assertKind(t, err, apperr.InputMissing)
}

func TestRejectsBadExtension(t *testing.T) {
	dir := t.TempDir()
	inp := filepath.Join(dir, "in.csv")
	if err := os.WriteFile(inp, []byte("dummy"), 0o644); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out})
	ae := assertKind(t, err, apperr.UnsupportedExtension)
	if ae.Error() != `unsupported extension "csv"; accepted: .xlsx, .xls` {
		t.Fatalf("message %q", ae.Error())
	}
}

func TestRejectsHeaderRowZero(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out, "--header-row", "0"})
	ae := assertKind(t, err, apperr.HeaderRowBelowOne)
	if ae.Error() != "--header-row must be >= 1, got 0" {
		t.Fatalf("message %q", ae.Error())
	}
}

func TestRejectsDuplicateColumns(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out, "--columns", "A,B,A"})
	assertKind(t, err, apperr.ColumnsDuplicate)
}

func TestRejectsEmptyColumnName(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out, "--columns", "A,,B"})
	assertKind(t, err, apperr.ColumnsEmpty)
}

func TestRejectsRenameEmpty(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out, "--rename", "A:"})
	assertKind(t, err, apperr.RenameEmptyName)
}

func TestRejectsRenameDestinationCollision(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out, "--rename", "A:X,B:X"})
	assertKind(t, err, apperr.RenameDestinationCollision)
}

func TestRejectsRenameDuplicateSource(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out, "--rename", "A:X,A:Y"})
	assertKind(t, err, apperr.RenameDuplicateSource)
}

func TestRejectsRenameBadPair(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	out := filepath.Join(dir, "out.json")
	_, err := parseArgs([]string{inp, out, "--rename", "A=B"})
	assertKind(t, err, apperr.RenameBadPair)
}

func TestRejectsOutputParentIsFile(t *testing.T) {
	dir := t.TempDir()
	inp := validInput(t, dir)
	blocker := filepath.Join(dir, "blocker")
	if err := os.WriteFile(blocker, []byte("i am a file"), 0o644); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(blocker, "out.json")
	_, err := parseArgs([]string{inp, out})
	assertKind(t, err, apperr.OutputParentNotDir)
}
