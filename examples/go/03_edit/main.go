//go:build office_oxide_dev

// 03_edit: edit documents in memory with ReplaceText and SetCell.
// Run: (cd examples/go/03_edit && go run -tags office_oxide_dev .)
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

func reportMarkdown(title string) string {
	return "# " + title + "\n\nRevenue grew in every region.\n\n" +
		"| Region | Q1  | Q2  |\n|--------|-----|-----|\n" +
		"| North  | 120 | 135 |\n| South  | 95  | 110 |\n"
}

// buildDocxBytes returns a DOCX. The Go binding has no DOCX writer: a DOCX
// comes from Markdown, which the library writes to a path, so this uses a
// private temp file for that hop and removes it before returning.
func buildDocxBytes(title string) []byte {
	f := must(os.CreateTemp("", "oo_example_*.docx"))
	path := f.Name()
	f.Close()
	defer os.Remove(path)
	check(oo.CreateFromMarkdown(reportMarkdown(title), "docx", path) == nil, "create the DOCX from Markdown")
	return must(os.ReadFile(path))
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

// textOf opens bytes as the given format and returns the plain text.
func textOf(data []byte, format string) string {
	doc := must(oo.OpenFromBytes(data, format))
	defer doc.Close()
	return must(doc.PlainText())
}

func main() {
	fmt.Println("step 1: build a DOCX in memory")
	docx := buildDocxBytes("Quarterly Report")

	fmt.Println("step 2: replace text")
	ed := must(oo.OpenEditableFromBytes(docx, "docx"))
	n := must(ed.ReplaceText("Quarterly Report", "Annual Report"))
	check(n == 1, "one replacement was made")
	fmt.Println("  replacements:", n)

	fmt.Println("step 3: save to bytes and check the text")
	out := must(ed.SaveToBytes())
	ed.Close()
	text := textOf(out, "docx")
	check(strings.Contains(text, "Annual Report"), "the new title is present")
	check(!strings.Contains(text, "Quarterly Report"), "the old title is gone")

	fmt.Println("step 4: build an XLSX in memory")
	xlsx := buildXlsxBytes()

	fmt.Println("step 5: set cells")
	ed = must(oo.OpenEditableFromBytes(xlsx, "xlsx"))
	check(ed.SetCell(0, "B2", oo.NewNumberCell(999)) == nil, "set B2")
	check(ed.SetCell(0, "A2", oo.NewStringCell("Central")) == nil, "set A2")

	fmt.Println("step 6: save to bytes and check the cells")
	out = must(ed.SaveToBytes())
	ed.Close()
	text = textOf(out, "xlsx")
	check(strings.Contains(text, "Central"), "A2 was replaced")
	check(strings.Contains(text, "999"), "B2 was replaced")
	check(!strings.Contains(text, "North"), "the old A2 is gone")

	fmt.Println("OK 03_edit")
}
