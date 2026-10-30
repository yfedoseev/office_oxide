//go:build office_oxide_dev

// 01_extract: read a document as plain text, Markdown and IR JSON.
// Run: (cd examples/go/01_extract && go run -tags office_oxide_dev .)
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

func main() {
	fmt.Println("step 1: build a DOCX in memory")
	data := buildDocxBytes("Quarterly Report")

	fmt.Println("step 2: open it from bytes")
	doc := openBytes(data, "docx")
	defer doc.Close()
	format := must(doc.Format())
	fmt.Println("  format:", format)
	check(format == "docx", "the opened format differs from the requested one")

	fmt.Println("step 3: extract plain text, Markdown and IR JSON")
	text := must(doc.PlainText())
	md := must(doc.ToMarkdown())
	ir := must(doc.ToIRJSON())
	fmt.Println("  chars:", len(text))

	fmt.Println("step 4: check the output")
	has(text, "Revenue grew in every region.")
	has(md, "Quarterly Report")
	check(strings.HasPrefix(ir, "{"), "the IR JSON is not an object")

	fmt.Println("OK 01_extract")
}
