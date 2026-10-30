//go:build office_oxide_dev

// 05_xlsx_writer: write a workbook with typed cells, a formula, a style, a merge and a column width, then read it back.
// Run: (cd examples/go/05_xlsx_writer && go run -tags office_oxide_dev .)
package main

import (
	"encoding/json"
	"fmt"
	"os"
	"strings"

	oo "github.com/yfedoseev/office_oxide/go"
)

// check fails the example loudly: it prints msg and exits non-zero.
func check(ok bool, msg string) {
	if !ok {
		fmt.Fprintln(os.Stderr, "FAIL:", msg)
		os.Exit(1)
	}
}

// must fails the example on an unexpected error.
func must[T any](v T, err error) T {
	if err != nil {
		fmt.Fprintln(os.Stderr, "FAIL:", err)
		os.Exit(1)
	}
	return v
}

// must0 fails the example on an unexpected error from a call with no result.
func must0(err error) { must(0, err) }

func has(text, expected string) {
	check(strings.Contains(text, expected), `"`+expected+`" is missing from the text`)
}

func openBytes(data []byte, format string) *oo.Document {
	return must(oo.OpenFromBytes(data, format))
}

// textOf opens bytes as the given format and returns the plain text.
func textOf(data []byte, format string) string {
	doc := openBytes(data, format)
	defer doc.Close()
	return must(doc.PlainText())
}

func buildXlsxBytes() []byte {
	w := oo.NewXlsxWriter()
	defer w.Close()
	s := w.AddSheet("Sales")
	for c, h := range []string{"Region", "Q1", "Q2", "Growth"} {
		must0(w.SetCellStyled(s, 0, uint32(c), h, true, "")) // bold header row
	}
	cells := []struct {
		row, col uint32
		v        any
	}{
		{1, 0, "North"}, {1, 1, 120}, {1, 2, 135}, {1, 3, true}, // D2
		{2, 0, "South"}, {2, 1, 95}, {2, 2, 110},
		{3, 0, "Total"}, {3, 1, oo.Formula("SUM(B2:B3)")}, {3, 2, oo.Formula("SUM(C2:C3)")},
		{5, 0, "Notes"},
	}
	for _, c := range cells {
		must0(w.SetCell(s, c.row, c.col, c.v))
	}
	w.MergeCells(s, 5, 0, 1, 3) // A6:C6
	w.SetColumnWidth(s, 0, 18)
	return must(w.ToBytes())
}

// sectionCount counts the sheets or slides in an IR JSON tree.
func sectionCount(irJSON string) int {
	var tree struct{ Sections []json.RawMessage }
	must0(json.Unmarshal([]byte(irJSON), &tree))
	return len(tree.Sections)
}

func main() {
	fmt.Println("step 1: build an XLSX in memory")
	data := buildXlsxBytes()

	fmt.Println("step 2: open it from bytes")
	doc := openBytes(data, "xlsx")
	defer doc.Close()
	text := must(doc.PlainText())
	sheets := sectionCount(must(doc.ToIRJSON()))
	fmt.Println("  format:", must(doc.Format()))
	fmt.Println("  chars:", len(text))
	fmt.Println("  sheets:", sheets)

	fmt.Println("step 3: check the cells")
	check(sheets == 1, "expected one sheet")
	for _, expected := range []string{"Sales", "Region", "Q1", "Q2", "Growth", "North", "120", "135", "South", "95", "110", "Total", "Notes"} {
		has(text, expected)
	}
	check(strings.Contains(text, "TRUE"), "the boolean cell is missing")
	check(strings.Contains(text, "SUM(B2:B3)") && strings.Contains(text, "SUM(C2:C3)"), "a formula is missing")

	fmt.Println("OK 05_xlsx_writer")
}
