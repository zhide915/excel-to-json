// Package apperr defines the typed errors shared by every stage; each Error()
// string is the exact single-line stderr message.
package apperr

import (
	"fmt"
	"strconv"
	"strings"
)

type Kind int

const (
	InputMissing Kind = iota
	InputNotFile
	UnsupportedExtension
	OutputParentNotDir
	HeaderRowBelowOne
	ColumnsEmpty
	ColumnsDuplicate
	RenameBadPair
	RenameEmptyName
	RenameDuplicateSource
	RenameDestinationCollision
	HeaderRowOutOfRange
	DuplicateHeaders
	ColumnsMissing
	RenameMissing
	RenameCollision
	NoData
	Parse
	Write
	XlsUnsupported
)

type Error struct {
	Kind Kind
	msg  string
}

func (e *Error) Error() string { return e.msg }

func Errf(kind Kind, format string, args ...any) *Error {
	return &Error{Kind: kind, msg: fmt.Sprintf(format, args...)}
}

// QuoteList renders a string slice as ["a", "b"] for error and verbose output.
func QuoteList(items []string) string {
	var b strings.Builder
	b.WriteByte('[')
	for i, s := range items {
		if i > 0 {
			b.WriteString(", ")
		}
		b.WriteString(strconv.Quote(s))
	}
	b.WriteByte(']')
	return b.String()
}
