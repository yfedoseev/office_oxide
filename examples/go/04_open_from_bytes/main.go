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

const markdown = `# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`

const png1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="

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

func buildPptxBytes() []byte {
	w := oo.NewPptxWriter()
	defer w.Close()
	w.SetPresentationSize(12192000, 6858000) // 16:9, in EMU (914400 per inch)
	one := w.AddSlide()
	must0(w.SetSlideTitle(one, "Welcome"))
	must0(w.AddSlideText(one, "Quarterly Report"))
	two := w.AddSlide()
	must0(w.SetSlideTitle(two, "Results"))
	must0(w.AddSlideText(two, "Revenue grew in every region."))
	png := must(base64.StdEncoding.DecodeString(png1x1))
	must0(w.AddSlideImage(two, png, "png", 914400, 2743200, 1828800, 1828800))
	return must(w.ToBytes())
}

func main() {
	fmt.Println("step 1: build a DOCX, an XLSX and a PPTX in memory")
	formats := []string{"docx", "xlsx", "pptx"}
	built := map[string][]byte{"docx": buildDocxBytes("Quarterly Report"), "xlsx": buildXlsxBytes(), "pptx": buildPptxBytes()}

	fmt.Println("step 2: open each from bytes")
	texts := map[string]string{}
	for _, format := range formats {
		doc := openBytes(built[format], format)
		got := must(doc.Format())
		check(got == format, "the opened format differs from the requested one")
		texts[format] = must(doc.PlainText())
		doc.Close()
		fmt.Println("  format:", format)
		fmt.Println("  chars:", len(texts[format]))
	}

	fmt.Println("step 3: check the text")
	has(texts["docx"], "Revenue grew in every region.")
	has(texts["xlsx"], "North")
	has(texts["xlsx"], "Total")
	has(texts["pptx"], "Welcome")
	has(texts["pptx"], "Results")

	fmt.Println("step 4: check that a wrong format is an error")
	_, err := oo.OpenFromBytes(built["xlsx"], "pptx")
	check(err != nil, "an XLSX opened as a PPTX was accepted")
	fmt.Println("  error:", err)

	fmt.Println("OK 04_open_from_bytes")
}
