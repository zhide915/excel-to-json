// Package cell holds the cell-value domain: the grid shape, Excel error
// markers, datetime formatting, date-format detection, and JSON scalar
// rendering. Pure: no I/O.
package cell

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
)

// Grid is a rectangular-ish sheet: rows of cells, each cell nil | bool |
// float64 | string.
type Grid = [][]any

// AllNull reports whether every cell in the row is null.
func AllNull(row []any) bool {
	for _, v := range row {
		if v != nil {
			return false
		}
	}
	return true
}

// Error markers that become JSON null: both native error cells and plain
// strings matching them.
var errorStrings = map[string]bool{
	"#REF!":         true,
	"#DIV/0!":       true,
	"#N/A":          true,
	"#VALUE!":       true,
	"#NAME?":        true,
	"#NULL!":        true,
	"#NUM!":         true,
	"#GETTING_DATA": true,
	"#SPILL!":       true,
	"#CALC!":        true,
}

func IsErrorString(s string) bool { return errorStrings[s] }

// FormatDateTime renders ISO 8601; date-only when the time is midnight.
func FormatDateTime(t time.Time) string {
	if t.Hour() == 0 && t.Minute() == 0 && t.Second() == 0 && t.Nanosecond() == 0 {
		return t.Format("2006-01-02")
	}
	return t.Format("2006-01-02T15:04:05")
}

// IsBuiltinDateNumFmt reports whether a builtin number-format ID renders as a
// date and/or time: 14-22 (dates/times), 27-36 and 50-58 (locale date formats),
// 45-47 (times).
func IsBuiltinDateNumFmt(id int) bool {
	switch {
	case id >= 14 && id <= 22,
		id >= 27 && id <= 36,
		id >= 45 && id <= 47,
		id >= 50 && id <= 58:
		return true
	}
	return false
}

// CustomFmtIsDateTime reports whether a custom number-format string renders as
// a date/time: it contains a y/m/d/h/s token outside quoted literals, escaped
// characters, and [] sections, except elapsed-time sections like [h] or [ss],
// which are time tokens themselves.
func CustomFmtIsDateTime(format string) bool {
	i := 0
	for i < len(format) {
		switch c := format[i]; c {
		case '"':
			i++
			for i < len(format) && format[i] != '"' {
				i++
			}
			i++
		case '\\', '_', '*': // escaped char / fill directives consume the next char
			i += 2
		case '[':
			j := i + 1
			for j < len(format) && format[j] != ']' {
				j++
			}
			section := strings.ToLower(format[i+1 : j])
			if section != "" && strings.Trim(section, "hms") == "" {
				return true
			}
			i = j + 1
		default:
			switch c | 0x20 { // ASCII lowercase
			case 'y', 'm', 'd', 'h', 's':
				return true
			}
			i++
		}
	}
	return false
}

// AppendJSONFloat appends f in encoding/json's canonical float format:
// shortest round-trip, 'f' notation for 1e-6 <= |f| < 1e21, else 'e' with a
// trimmed exponent.
func AppendJSONFloat(buf []byte, f float64) []byte {
	abs := math.Abs(f)
	format := byte('f')
	if abs != 0 && (abs < 1e-6 || abs >= 1e21) {
		format = 'e'
	}
	buf = strconv.AppendFloat(buf, f, format, -1, 64)
	if format == 'e' {
		// Trim leading zero from a single-digit exponent: 2.5e-09 -> 2.5e-9.
		if n := len(buf); n >= 4 && buf[n-4] == 'e' && buf[n-3] == '-' && buf[n-2] == '0' {
			buf[n-2] = buf[n-1]
			buf = buf[:n-1]
		}
	}
	return buf
}

// AppendJSONString appends s as a JSON string with minimal escaping: only ",
// \, and control characters; all other Unicode passes through as raw UTF-8.
func AppendJSONString(buf []byte, s string) []byte {
	buf = append(buf, '"')
	start := 0
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c != '"' && c != '\\' && c >= 0x20 {
			continue
		}
		buf = append(buf, s[start:i]...)
		switch c {
		case '"':
			buf = append(buf, '\\', '"')
		case '\\':
			buf = append(buf, '\\', '\\')
		case '\n':
			buf = append(buf, '\\', 'n')
		case '\r':
			buf = append(buf, '\\', 'r')
		case '\t':
			buf = append(buf, '\\', 't')
		case '\b':
			buf = append(buf, '\\', 'b')
		case '\f':
			buf = append(buf, '\\', 'f')
		default:
			buf = append(buf, fmt.Sprintf(`\u%04x`, c)...)
		}
		start = i + 1
	}
	buf = append(buf, s[start:]...)
	return append(buf, '"')
}

// AppendJSONValue appends one cell value; non-cell types fall back to their
// string form.
func AppendJSONValue(buf []byte, v any) []byte {
	switch t := v.(type) {
	case nil:
		return append(buf, "null"...)
	case bool:
		if t {
			return append(buf, "true"...)
		}
		return append(buf, "false"...)
	case float64:
		if math.IsNaN(t) || math.IsInf(t, 0) {
			return append(buf, "null"...)
		}
		return AppendJSONFloat(buf, t)
	case string:
		return AppendJSONString(buf, t)
	default:
		return AppendJSONString(buf, fmt.Sprint(t))
	}
}
