//go:build office_oxide_dev

// 05_xlsx_writer: write a workbook with typed cells, formulas, styling, a merge and a column width.
// Run: (cd examples/go/05_xlsx_writer && go run -tags office_oxide_dev .)
package main

import (
	"fmt"
	"os"
	"strings"

	oo "github.com/yfedoseev/office_oxide/go"
)

// check fails the example loudly: it prints msg and exits non-zero.
func check(ok bool, msg string) {
	if !ok {
		fmt.Fprintln(os.Stderr, "FAILED:", msg)
		os.Exit(1)
	}
}

// must fails the example on an unexpected error.
func must[T any](v T, err error) T {
	if err != nil {
		fmt.Fprintln(os.Stderr, "FAILED:", err)
		os.Exit(1)
	}
	return v
}

// buildXlsxBytes returns a workbook with one sheet, "Sales": headers (bold on
// grey), two data rows, a SUM row, a boolean, a merged note and a wide column A.
func buildXlsxBytes() []byte {
	w := oo.NewXlsxWriter()
	defer w.Close()
	s := w.AddSheet("Sales")
	for c, h := range []string{"Region", "Q1", "Q2", "Growth"} {
		check(w.SetCellStyled(s, 0, uint32(c), h, true, "D3D3D3") == nil, "write a header cell")
	}
	cells := []struct {
		row, col uint32
		v        any
	}{
		{1, 0, "North"}, {1, 1, 120}, {1, 2, 135}, {1, 3, true},
		{2, 0, "South"}, {2, 1, 95}, {2, 2, 110},
		{3, 0, "Total"}, {3, 1, oo.Formula("=SUM(B2:B3)")}, {3, 2, oo.Formula("=SUM(C2:C3)")},
		{5, 0, "Notes"},
	}
	for _, c := range cells {
		check(w.SetCell(s, c.row, c.col, c.v) == nil, "write a cell")
	}
	w.MergeCells(s, 5, 0, 1, 3)
	w.SetColumnWidth(s, 0, 18)
	return must(w.ToBytes())
}

func main() {
	fmt.Println("step 1: build an XLSX in memory")
	data := buildXlsxBytes() // see above for the cell-by-cell calls

	fmt.Println("step 2: check that a write to a missing sheet is refused")
	w := oo.NewXlsxWriter()
	defer w.Close()
	check(w.SetCell(9, 0, 0, "x") != nil, "a missing sheet is an error, not a silent drop")

	fmt.Println("step 3: open it from bytes")
	doc := must(oo.OpenFromBytes(data, "xlsx"))
	defer doc.Close()
	text := must(doc.PlainText())
	format := must(doc.Format())
	fmt.Println("  format:", format)
	fmt.Println("  chars:", len(text))
	fmt.Println("  sheets: 1")

	fmt.Println("step 4: check the cells")
	check(strings.Contains(text, "Sales"), "the sheet name is present")
	check(strings.Contains(text, "North"), "text cells are present")
	check(strings.Contains(text, "135"), "number cells are present")
	check(strings.Contains(text, "TRUE"), "the boolean cell is present")
	check(strings.Contains(text, "=SUM(B2:B3)"), "the formula survived")
	check(strings.Contains(text, "Notes"), "the merged cell text is present")

	fmt.Println("OK 05_xlsx_writer")
}
