//go:build office_oxide_dev

// 03_edit: replace text in a DOCX and set cells in an XLSX, in memory.
// Run: (cd examples/go/03_edit && go run -tags office_oxide_dev .)
package main

import (
	"fmt"
	"os"
	"strings"

	oo "github.com/yfedoseev/office_oxide/go"
)

const markdown = `# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`

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

// markdownToBytes creates a document from Markdown. The library writes to a
// path, so the file lives in a private temp file just long enough to read its
// bytes back.
func markdownToBytes(md, format string) []byte {
	f := must(os.CreateTemp("", "oo_example_*."+format))
	path := f.Name()
	f.Close()
	defer os.Remove(path)
	must0(oo.CreateFromMarkdown(md, format, path))
	return must(os.ReadFile(path))
}

func buildDocxBytes(title string) []byte {
	return markdownToBytes(strings.ReplaceAll(markdown, "Quarterly Report", title), "docx")
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

func main() {
	fmt.Println("step 1: build a DOCX in memory")
	docx := buildDocxBytes("Quarterly Report")

	fmt.Println("step 2: replace text")
	ed := must(oo.OpenEditableFromBytes(docx, "docx"))
	replaced := must(ed.ReplaceText("Quarterly Report", "Annual Report"))
	fmt.Println("  replacements:", replaced)
	check(replaced == 1, "replace_text did not make exactly one replacement")

	fmt.Println("step 3: save to bytes and check the text")
	edited := must(ed.SaveToBytes())
	ed.Close()
	text := textOf(edited, "docx")
	has(text, "Annual Report")
	check(!strings.Contains(text, "Quarterly Report"), "the old title is still present")

	fmt.Println("step 4: build an XLSX in memory")
	xlsx := buildXlsxBytes()

	fmt.Println("step 5: set cells")
	ed = must(oo.OpenEditableFromBytes(xlsx, "xlsx"))
	must0(ed.SetCell(0, "B2", oo.NewNumberCell(999)))
	must0(ed.SetCell(0, "A2", oo.NewStringCell("Central")))

	fmt.Println("step 6: save to bytes and check the cells")
	edited = must(ed.SaveToBytes())
	ed.Close()
	text = textOf(edited, "xlsx")
	has(text, "Central")
	has(text, "999")
	check(!strings.Contains(text, "North"), "the old A2 value is still present")

	fmt.Println("OK 03_edit")
}
