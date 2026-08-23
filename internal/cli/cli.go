// Package cli owns argument parsing, pre-read validation, pipeline
// orchestration, and process exit codes.
package cli

import (
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"slices"
	"sort"
	"strconv"
	"strings"

	"github.com/zhide915/excel-to-json/internal/apperr"
	"github.com/zhide915/excel-to-json/internal/pipeline"
	"github.com/zhide915/excel-to-json/internal/reader"
)

const version = "1.0.0"

const usage = `Convert an Excel spreadsheet (.xlsx/.xls) to compact JSON.

Usage: excel-to-json <INPUT> <OUTPUT> [OPTIONS]

Arguments:
  <INPUT>   Input Excel file (.xlsx or .xls)
  <OUTPUT>  Output JSON file

Options:
      --header-row <N>       1-based row containing headers [default: 1]
      --columns <LIST>       Comma-separated list of columns to keep
      --rename <PAIRS>       Comma-separated rename pairs: old1:new1,old2:new2
      --ignore-empty         Drop rows and columns that are entirely null
      --verbose              Print pipeline narrative to stderr
      --help                 Print help
      --version              Print version

Example:
  excel-to-json payroll.xlsx out.json --header-row 2 \
    --columns 'EmpID,Salary' --rename 'EmpID:employee_id' --ignore-empty
`

// Main is the whole program: parse, validate, run. Errors go to stderr as
// "error: ..." with exit 1; --help/--version exit 0.
func Main(argv []string) int {
	args, err := parseArgs(argv)
	if err != nil {
		var uerr *usageError
		switch {
		case errors.Is(err, flag.ErrHelp):
			fmt.Print(usage)
			return 0
		case errors.Is(err, errVersionRequested):
			fmt.Println("excel-to-json " + version)
			return 0
		case errors.As(err, &uerr):
			fmt.Fprintf(os.Stderr, "error: %s\n\n%s", err, usage)
			return 1
		default:
			fmt.Fprintf(os.Stderr, "error: %s\n", err)
			return 1
		}
	}

	if err := run(args); err != nil {
		fmt.Fprintf(os.Stderr, "error: %s\n", err)
		return 1
	}
	return 0
}

// errVersionRequested signals --version; Main prints the version and exits 0.
var errVersionRequested = errors.New("version requested")

// usageError wraps a flag/positional parse error; Main appends the usage text.
type usageError struct{ err error }

func (u *usageError) Error() string { return u.err.Error() }

type cliArgs struct {
	input       string
	output      string
	headerRow   int
	columns     []string              // nil when --columns is not set
	rename      []pipeline.RenamePair // nil when --rename is not set
	ignoreEmpty bool
	verbose     bool
}

// parseArgs parses argv (without the program name) and runs every pre-read
// validation. Options may follow the positionals, which stdlib flag stops at,
// so re-parse after pulling off each positional.
func parseArgs(argv []string) (*cliArgs, error) {
	fs := flag.NewFlagSet("excel-to-json", flag.ContinueOnError)
	fs.SetOutput(io.Discard)
	headerRow := fs.Int64("header-row", 1, "")
	columns := fs.String("columns", "", "")
	rename := fs.String("rename", "", "")
	ignoreEmpty := fs.Bool("ignore-empty", false, "")
	verbose := fs.Bool("verbose", false, "")
	showVersion := fs.Bool("version", false, "")

	var positionals []string
	rest := argv
	for len(rest) > 0 {
		if err := fs.Parse(rest); err != nil {
			if errors.Is(err, flag.ErrHelp) {
				return nil, flag.ErrHelp
			}
			return nil, &usageError{err}
		}
		if fs.NArg() == 0 {
			break
		}
		positionals = append(positionals, fs.Arg(0))
		rest = fs.Args()[1:]
	}

	if *showVersion {
		return nil, errVersionRequested
	}
	switch {
	case len(positionals) < 2:
		missing := []string{"<INPUT>", "<OUTPUT>"}[len(positionals):]
		return nil, &usageError{fmt.Errorf("missing required argument(s): %s", strings.Join(missing, " "))}
	case len(positionals) > 2:
		return nil, &usageError{fmt.Errorf("unexpected argument %s", strconv.Quote(positionals[2]))}
	}

	if *headerRow < 1 {
		return nil, apperr.Errf(apperr.HeaderRowBelowOne, "--header-row must be >= 1, got %d", *headerRow)
	}

	provided := map[string]bool{}
	fs.Visit(func(f *flag.Flag) { provided[f.Name] = true })

	args := &cliArgs{
		input:       positionals[0],
		output:      positionals[1],
		headerRow:   int(*headerRow),
		ignoreEmpty: *ignoreEmpty,
		verbose:     *verbose,
	}
	if provided["columns"] {
		cols, err := parseColumns(*columns)
		if err != nil {
			return nil, err
		}
		args.columns = cols
	}
	if provided["rename"] {
		pairs, err := parseRename(*rename)
		if err != nil {
			return nil, err
		}
		args.rename = pairs
	}

	if err := validateInputPath(args.input); err != nil {
		return nil, err
	}
	if err := validateOutputPath(args.output); err != nil {
		return nil, err
	}
	return args, nil
}

func parseColumns(raw string) ([]string, error) {
	items := strings.Split(raw, ",")
	for i := range items {
		items[i] = strings.TrimSpace(items[i])
	}
	if slices.Contains(items, "") {
		return nil, apperr.Errf(apperr.ColumnsEmpty, "--columns contains an empty name")
	}
	if dupes := pipeline.Duplicates(items); len(dupes) > 0 {
		return nil, apperr.Errf(apperr.ColumnsDuplicate, "--columns has duplicate entries: %s", apperr.QuoteList(dupes))
	}
	return items, nil
}

func parseRename(raw string) ([]pipeline.RenamePair, error) {
	var pairs []pipeline.RenamePair
	seenSources := map[string]bool{}
	for _, piece := range strings.Split(raw, ",") {
		oldName, newName, ok := strings.Cut(piece, ":")
		if !ok {
			return nil, apperr.Errf(apperr.RenameBadPair, "--rename entry %s is not in 'old:new' form", strconv.Quote(piece))
		}
		oldName = strings.TrimSpace(oldName)
		newName = strings.TrimSpace(newName)
		if oldName == "" || newName == "" {
			return nil, apperr.Errf(apperr.RenameEmptyName, "--rename entry has an empty name: %s", strconv.Quote(piece))
		}
		if seenSources[oldName] {
			return nil, apperr.Errf(apperr.RenameDuplicateSource, "--rename has a duplicate source key: %s", strconv.Quote(oldName))
		}
		seenSources[oldName] = true
		pairs = append(pairs, pipeline.RenamePair{Old: oldName, New: newName})
	}

	destSeen := map[string]bool{}
	var collisions []string
	for _, p := range pairs {
		if destSeen[p.New] {
			collisions = append(collisions, p.New)
		}
		destSeen[p.New] = true
	}
	if len(collisions) > 0 {
		sort.Strings(collisions)
		collisions = slices.Compact(collisions)
		return nil, apperr.Errf(apperr.RenameDestinationCollision, "--rename produces destination-name collisions: %s", apperr.QuoteList(collisions))
	}
	return pairs, nil
}

func validateInputPath(path string) error {
	info, err := os.Stat(path)
	if err != nil {
		return apperr.Errf(apperr.InputMissing, "input file does not exist: %s", path)
	}
	if !info.Mode().IsRegular() {
		return apperr.Errf(apperr.InputNotFile, "input path is not a file: %s", path)
	}
	ext := strings.TrimPrefix(strings.ToLower(filepath.Ext(path)), ".")
	if ext != "xlsx" && ext != "xls" {
		return apperr.Errf(apperr.UnsupportedExtension, "unsupported extension %s; accepted: .xlsx, .xls", strconv.Quote(ext))
	}
	return nil
}

// validateOutputPath requires only that the immediate parent, if it exists, is
// a directory; missing parents are created at write time.
func validateOutputPath(path string) error {
	parent := filepath.Dir(path)
	if parent == "" || parent == "." {
		return nil
	}
	if info, err := os.Stat(parent); err == nil && !info.IsDir() {
		return apperr.Errf(apperr.OutputParentNotDir, "output parent path is not a directory: %s", parent)
	}
	return nil
}

// run chains the pipeline stages in order, narrating each step to stderr under
// --verbose.
func run(a *cliArgs) error {
	g, err := reader.ReadExcel(a.input)
	if err != nil {
		return err
	}
	cols := 0
	if len(g) > 0 {
		cols = len(g[0])
	}
	vprintf(a.verbose, "[1/5] Read sheet: %d rows × %d columns", len(g), cols)

	headers, data, err := pipeline.ExtractHeaders(g, a.headerRow)
	if err != nil {
		return err
	}
	vprintf(a.verbose, "[2/5] Header row %d: %s\n      Data rows remaining: %d",
		a.headerRow, apperr.QuoteList(headers), len(data))

	if a.columns != nil {
		headers, data, err = pipeline.FilterColumns(headers, data, a.columns)
		if err != nil {
			return err
		}
		vprintf(a.verbose, "[3/5] --columns filter: kept %d columns %s", len(headers), apperr.QuoteList(headers))
	} else {
		vprintf(a.verbose, "[3/5] --columns: (skipped)")
	}

	if a.rename != nil {
		headers, data, err = pipeline.RenameColumns(headers, data, a.rename)
		if err != nil {
			return err
		}
		vprintf(a.verbose, "[4/5] --rename: %s", apperr.QuoteList(headers))
	} else {
		vprintf(a.verbose, "[4/5] --rename: (skipped)")
	}

	if a.ignoreEmpty {
		beforeRows, beforeCols := len(data), len(headers)
		headers, data = pipeline.DropEmpty(headers, data, true)
		vprintf(a.verbose, "[5/5] --ignore-empty: dropped %d empty rows, %d empty columns. %d records remain.",
			beforeRows-len(data), beforeCols-len(headers), len(data))
	} else {
		vprintf(a.verbose, "[5/5] --ignore-empty: (skipped)")
	}

	if err := pipeline.WriteJSON(headers, data, a.output); err != nil {
		return err
	}
	vprintf(a.verbose, "Wrote %d records to %s", len(data), a.output)
	return nil
}

func vprintf(enabled bool, format string, args ...any) {
	if enabled {
		fmt.Fprintf(os.Stderr, format+"\n", args...)
	}
}
