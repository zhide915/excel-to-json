package cell

import (
	"math"
	"testing"
	"time"
)

func jsonValue(v any) string { return string(AppendJSONValue(nil, v)) }

func TestNilIsNull(t *testing.T) {
	if got := jsonValue(nil); got != "null" {
		t.Fatalf("got %q", got)
	}
}

func TestBoolValues(t *testing.T) {
	if got := jsonValue(true); got != "true" {
		t.Fatalf("got %q", got)
	}
	if got := jsonValue(false); got != "false" {
		t.Fatalf("got %q", got)
	}
}

func TestIntegralFloatValue(t *testing.T) {
	if got := jsonValue(42.0); got != "42" {
		t.Fatalf("got %q", got)
	}
}

func TestFloatValue(t *testing.T) {
	if got := jsonValue(2.5); got != "2.5" {
		t.Fatalf("got %q", got)
	}
}

func TestFloatNaNIsNull(t *testing.T) {
	if got := jsonValue(math.NaN()); got != "null" {
		t.Fatalf("got %q", got)
	}
}

func TestFloatInfinityIsNull(t *testing.T) {
	if got := jsonValue(math.Inf(1)); got != "null" {
		t.Fatalf("got %q", got)
	}
	if got := jsonValue(math.Inf(-1)); got != "null" {
		t.Fatalf("got %q", got)
	}
}

func TestFloatScientificNotation(t *testing.T) {
	cases := map[float64]string{
		1e21:   "1e+21",
		2.5e-9: "2.5e-9",
		0:      "0",
		-1.5:   "-1.5",
	}
	for in, want := range cases {
		if got := jsonValue(in); got != want {
			t.Errorf("float %v: got %q, want %q", in, got, want)
		}
	}
}

func TestStringPassthrough(t *testing.T) {
	if got := jsonValue("hello"); got != `"hello"` {
		t.Fatalf("got %q", got)
	}
}

func TestEmptyStringPreserved(t *testing.T) {
	if got := jsonValue(""); got != `""` {
		t.Fatalf("got %q", got)
	}
}

func TestStringEscaping(t *testing.T) {
	cases := map[string]string{
		"a\"b":    `"a\"b"`,
		`a\b`:     `"a\\b"`,
		"a\nb":    `"a\nb"`,
		"a\tb":    `"a\tb"`,
		"<html>&": `"<html>&"`, // no HTML escaping
	}
	for in, want := range cases {
		if got := jsonValue(in); got != want {
			t.Errorf("string %q: got %q, want %q", in, got, want)
		}
	}
}

func TestControlCharEscaped(t *testing.T) {
	in := "a\x01b"
	want := `"a` + `\` + `u0001b"` // concatenated so the escape stays literal text
	if got := jsonValue(in); got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestExcelErrorStrings(t *testing.T) {
	for _, s := range []string{"#REF!", "#DIV/0!", "#N/A", "#VALUE!", "#NAME?",
		"#NULL!", "#NUM!", "#GETTING_DATA", "#SPILL!", "#CALC!"} {
		if !IsErrorString(s) {
			t.Errorf("%q should be an error marker", s)
		}
	}
	if IsErrorString("#NOPE!") || IsErrorString("plain") {
		t.Error("non-markers misclassified")
	}
}

func TestDateOnlyFormatting(t *testing.T) {
	d := time.Date(2026, 4, 16, 0, 0, 0, 0, time.UTC)
	if got := FormatDateTime(d); got != "2026-04-16" {
		t.Fatalf("got %q", got)
	}
}

func TestDateTimeFormatting(t *testing.T) {
	dt := time.Date(2026, 4, 16, 14, 30, 0, 0, time.UTC)
	if got := FormatDateTime(dt); got != "2026-04-16T14:30:00" {
		t.Fatalf("got %q", got)
	}
}

func TestBuiltinDateNumFmt(t *testing.T) {
	for _, id := range []int{14, 15, 22, 27, 36, 45, 47, 50, 58} {
		if !IsBuiltinDateNumFmt(id) {
			t.Errorf("id %d should be a date format", id)
		}
	}
	for _, id := range []int{0, 1, 2, 9, 13, 23, 26, 37, 44, 48, 49, 59} {
		if IsBuiltinDateNumFmt(id) {
			t.Errorf("id %d should NOT be a date format", id)
		}
	}
}

func TestCustomFmtDateDetection(t *testing.T) {
	dateFormats := []string{"yyyy-mm-dd", "d-mmm-yy", "hh:mm:ss", "[h]:mm", "m/d/yy h:mm", "yyyy\"年\"m\"月\""}
	for _, f := range dateFormats {
		if !CustomFmtIsDateTime(f) {
			t.Errorf("%q should be detected as date/time", f)
		}
	}
	nonDateFormats := []string{"General", "#,##0.00", "0.00%", "0.00E+00", "[Red]#,##0", `"months" 0.00`, "@", "# ?/?"}
	for _, f := range nonDateFormats {
		if CustomFmtIsDateTime(f) {
			t.Errorf("%q should NOT be detected as date/time", f)
		}
	}
}

func TestUnicodePreserved(t *testing.T) {
	v := "测试 日本語 한국어"
	if got := jsonValue(v); got != `"`+v+`"` {
		t.Fatalf("got %q", got)
	}
}

func TestAllNull(t *testing.T) {
	if !AllNull([]any{nil, nil}) || !AllNull(nil) {
		t.Error("all-nil rows should report true")
	}
	if AllNull([]any{nil, false}) {
		t.Error("row with a value should report false")
	}
}
