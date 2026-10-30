//go:build office_oxide_dev

// 04_open_from_bytes: build a DOCX, an XLSX and a PPTX in memory and open each from its bytes.
// Run: (cd examples/go/04_open_from_bytes && go run -tags office_oxide_dev .)
package main

import (
	"encoding/base64"
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

const png1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="

// buildPptxBytes returns two slides: "Welcome" and "Results" (with an image).
func buildPptxBytes() []byte {
	w := oo.NewPptxWriter()
	defer w.Close()
	w.SetPresentationSize(9144000, 6858000) // 4:3; 914400 EMU = 1 inch
	s1 := w.AddSlide()
	check(w.SetSlideTitle(s1, "Welcome") == nil, "set the slide 1 title")
	check(w.AddSlideText(s1, "Quarterly Report") == nil, "add the slide 1 text")
	s2 := w.AddSlide()
	check(w.SetSlideTitle(s2, "Results") == nil, "set the slide 2 title")
	check(w.AddSlideText(s2, "Revenue grew in every region.") == nil, "add the slide 2 text")
	png := must(base64.StdEncoding.DecodeString(png1x1))
	check(w.AddSlideImage(s2, png, "png", 914400, 1371600, 914400, 914400) == nil, "add the image")
	return must(w.ToBytes())
}

// textOf opens bytes as the given format and returns the plain text.
func textOf(data []byte, format string) string {
	doc := must(oo.OpenFromBytes(data, format))
	defer doc.Close()
	return must(doc.PlainText())
}

func main() {
	fmt.Println("step 1: build a DOCX, an XLSX and a PPTX in memory")
	docx, xlsx, pptx := buildDocxBytes("Quarterly Report"), buildXlsxBytes(), buildPptxBytes()

	fmt.Println("step 2: open each from bytes")
	dt, xt, pt := textOf(docx, "docx"), textOf(xlsx, "xlsx"), textOf(pptx, "pptx")
	for _, r := range []struct{ f, t string }{{"docx", dt}, {"xlsx", xt}, {"pptx", pt}} {
		fmt.Println("  format:", r.f)
		fmt.Println("  chars:", len(r.t))
	}

	fmt.Println("step 3: check the text")
	check(strings.Contains(dt, "Revenue grew in every region."), "the DOCX has its paragraph")
	check(strings.Contains(xt, "North"), "the XLSX has its first data row")
	check(strings.Contains(pt, "Welcome"), "the PPTX has its first title")

	fmt.Println("step 4: check that a wrong format is an error")
	_, err := oo.OpenFromBytes(xlsx, "pptx")
	check(err != nil, "an XLSX opened as a PPTX is refused")
	fmt.Println("  error:", err)

	fmt.Println("OK 04_open_from_bytes")
}
