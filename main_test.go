// End-to-end tests: build the real binary once, invoke it via os/exec.
package main

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

const sheet1 = "Sheet1"

var binPath string

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

func TestMain(m *testing.M) {
	dir, err := os.MkdirTemp("", "excel-to-json-bin")
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	binPath = filepath.Join(dir, "excel-to-json")
	if runtime.GOOS == "windows" {
		binPath += ".exe"
	}
	out, err := exec.Command("go", "build", "-o", binPath, ".").CombinedOutput()
	if err != nil {
		fmt.Fprintf(os.Stderr, "go build failed: %v\n%s", err, out)
		os.Exit(1)
	}
	code := m.Run()
	os.RemoveAll(dir)
	os.Exit(code)
}

type runResult struct {
	exitCode int
	stdout   string
	stderr   string
}

func runBinary(t *testing.T, args ...string) runResult {
	t.Helper()
	cmd := exec.Command(binPath, args...)
	var stdout, stderr strings.Builder
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	err := cmd.Run()
	code := 0
	if err != nil {
		ee, ok := err.(*exec.ExitError)
		if !ok {
			t.Fatalf("failed to run binary: %v", err)
		}
		code = ee.ExitCode()
	}
	return runResult{exitCode: code, stdout: stdout.String(), stderr: stderr.String()}
}

func makeSimpleXlsx(t *testing.T, dir, name string) string {
	t.Helper()
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "EmpID")
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

func TestHappyPath(t *testing.T) {
	dir := t.TempDir()
	inp := makeSimpleXlsx(t, dir, "in")
	out := filepath.Join(dir, "out.json")

	res := runBinary(t, inp, out)
	if res.exitCode != 0 {
		t.Fatalf("exit %d, stderr: %s", res.exitCode, res.stderr)
	}
	text, err := os.ReadFile(out)
	if err != nil {
		t.Fatal(err)
	}
	want := `[{"EmpID":1,"Name":"Alice","Salary":5000},{"EmpID":2,"Name":"Bob","Salary":6000}]`
	if string(text) != want {
		t.Fatalf("got %q", text)
	}
}

func TestFilterRenameAndIgnoreEmpty(t *testing.T) {
	dir := t.TempDir()
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "Payroll Q1")
	f.SetCellStr(sheet1, "A3", "EmpID")
	f.SetCellStr(sheet1, "B3", "Name")
	f.SetCellStr(sheet1, "C3", "Salary")
	f.SetCellStr(sheet1, "D3", "InternalNote")
	f.SetCellValue(sheet1, "A4", 1)
	f.SetCellStr(sheet1, "B4", "Alice")
	f.SetCellValue(sheet1, "C4", 5000)
	f.SetCellStr(sheet1, "D4", "VIP")
	f.SetCellValue(sheet1, "A5", 2)
	f.SetCellStr(sheet1, "B5", "Bob")
	f.SetCellValue(sheet1, "C5", 6000)
	inp := saveWorkbook(t, f, dir, "full.xlsx")
	out := filepath.Join(dir, "out.json")

	res := runBinary(t, inp, out,
		"--header-row", "3",
		"--columns", "EmpID,Name,Salary",
		"--rename", "EmpID:employee_id,Name:full_name,Salary:monthly_salary",
		"--ignore-empty",
	)
	if res.exitCode != 0 {
		t.Fatalf("exit %d, stderr: %s", res.exitCode, res.stderr)
	}
	text, err := os.ReadFile(out)
	if err != nil {
		t.Fatal(err)
	}
	want := `[{"employee_id":1,"full_name":"Alice","monthly_salary":5000},{"employee_id":2,"full_name":"Bob","monthly_salary":6000}]`
	if string(text) != want {
		t.Fatalf("got %q", text)
	}
}

func TestDeterminismByteIdentical(t *testing.T) {
	dir := t.TempDir()
	inp := makeSimpleXlsx(t, dir, "det")
	out1 := filepath.Join(dir, "o1.json")
	out2 := filepath.Join(dir, "o2.json")

	if res := runBinary(t, inp, out1); res.exitCode != 0 {
		t.Fatalf("run 1: exit %d, stderr: %s", res.exitCode, res.stderr)
	}
	if res := runBinary(t, inp, out2); res.exitCode != 0 {
		t.Fatalf("run 2: exit %d, stderr: %s", res.exitCode, res.stderr)
	}

	b1, _ := os.ReadFile(out1)
	b2, _ := os.ReadFile(out2)
	if string(b1) != string(b2) {
		t.Fatal("outputs differ between runs")
	}
}

func TestUnicodeRoundTrip(t *testing.T) {
	dir := t.TempDir()
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "名前")
	f.SetCellStr(sheet1, "B1", "国")
	f.SetCellStr(sheet1, "A2", "田中")
	f.SetCellStr(sheet1, "B2", "日本")
	f.SetCellStr(sheet1, "A3", "김철수")
	f.SetCellStr(sheet1, "B3", "한국")
	inp := saveWorkbook(t, f, dir, "u.xlsx")
	out := filepath.Join(dir, "out.json")

	if res := runBinary(t, inp, out); res.exitCode != 0 {
		t.Fatalf("exit %d, stderr: %s", res.exitCode, res.stderr)
	}
	text, _ := os.ReadFile(out)
	s := string(text)
	if !strings.Contains(s, "田中") || !strings.Contains(s, "한국") || strings.Contains(s, `\u`) {
		t.Fatalf("got %q", s)
	}
}

func TestMissingColumnErrorsNonzero(t *testing.T) {
	dir := t.TempDir()
	inp := makeSimpleXlsx(t, dir, "e")
	out := filepath.Join(dir, "out.json")

	res := runBinary(t, inp, out, "--columns", "EmpID,DoesNotExist")
	if res.exitCode != 1 {
		t.Fatalf("exit %d", res.exitCode)
	}
	if !strings.Contains(res.stderr, "DoesNotExist") {
		t.Fatalf("stderr %q", res.stderr)
	}
	if _, err := os.Stat(out); !os.IsNotExist(err) {
		t.Fatal("output file should not exist")
	}
}

func TestQuietByDefault(t *testing.T) {
	dir := t.TempDir()
	inp := makeSimpleXlsx(t, dir, "q")
	out := filepath.Join(dir, "out.json")

	res := runBinary(t, inp, out)
	if res.exitCode != 0 {
		t.Fatalf("exit %d, stderr: %s", res.exitCode, res.stderr)
	}
	if res.stdout != "" || res.stderr != "" {
		t.Fatalf("stdout %q, stderr %q", res.stdout, res.stderr)
	}
}

func TestVerboseEmitsStepPrefixes(t *testing.T) {
	dir := t.TempDir()
	inp := makeSimpleXlsx(t, dir, "v")
	out := filepath.Join(dir, "out.json")

	res := runBinary(t, inp, out, "--verbose")
	if res.exitCode != 0 {
		t.Fatalf("exit %d, stderr: %s", res.exitCode, res.stderr)
	}
	for _, want := range []string{"[1/5] Read sheet", "[2/5] Header row 1", "Wrote 2 records"} {
		if !strings.Contains(res.stderr, want) {
			t.Errorf("stderr missing %q:\n%s", want, res.stderr)
		}
	}
	if res.stdout != "" {
		t.Fatalf("stdout should be empty, got %q", res.stdout)
	}
}

func TestHelpExitsZero(t *testing.T) {
	res := runBinary(t, "--help")
	if res.exitCode != 0 {
		t.Fatalf("exit %d", res.exitCode)
	}
	if !strings.Contains(res.stdout, "excel-to-json") || !strings.Contains(res.stdout, "--header-row") {
		t.Fatalf("stdout %q", res.stdout)
	}
}

func TestOutputIsValidJSON(t *testing.T) {
	dir := t.TempDir()
	inp := makeSimpleXlsx(t, dir, "j")
	out := filepath.Join(dir, "out.json")

	if res := runBinary(t, inp, out); res.exitCode != 0 {
		t.Fatalf("exit %d, stderr: %s", res.exitCode, res.stderr)
	}
	text, _ := os.ReadFile(out)
	var parsed []map[string]any
	if err := json.Unmarshal(text, &parsed); err != nil {
		t.Fatalf("not valid JSON: %v", err)
	}
}

// Slow; skipped by `go test -short`.
func TestLargeFile10kRows(t *testing.T) {
	if testing.Short() {
		t.Skip("skipping 10k-row performance test in -short mode")
	}
	dir := t.TempDir()
	f := excelize.NewFile()
	f.SetCellStr(sheet1, "A1", "id")
	f.SetCellStr(sheet1, "B1", "name")
	f.SetCellStr(sheet1, "C1", "value")
	for i := 0; i < 10000; i++ {
		row := i + 2
		f.SetCellValue(sheet1, fmt.Sprintf("A%d", row), i)
		f.SetCellStr(sheet1, fmt.Sprintf("B%d", row), fmt.Sprintf("user_%d", i))
		f.SetCellValue(sheet1, fmt.Sprintf("C%d", row), float64(i)*1.5)
	}
	inp := saveWorkbook(t, f, dir, "big.xlsx")
	out := filepath.Join(dir, "out.json")

	if res := runBinary(t, inp, out); res.exitCode != 0 {
		t.Fatalf("exit %d, stderr: %s", res.exitCode, res.stderr)
	}
	text, _ := os.ReadFile(out)
	var parsed []map[string]any
	if err := json.Unmarshal(text, &parsed); err != nil {
		t.Fatal(err)
	}
	if len(parsed) != 10000 {
		t.Fatalf("got %d records", len(parsed))
	}
}
