//go:build office_oxide_dev

// 02_create_from_markdown: turn one Markdown source into a DOCX, an XLSX and a PPTX.
// Run: (cd examples/go/02_create_from_markdown && go run -tags office_oxide_dev .)
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
	fmt.Println("step 1: write the Markdown source")

	fmt.Println("step 2: create a DOCX, an XLSX and a PPTX from it")
	formats := []string{"docx", "xlsx", "pptx"}
	created := map[string][]byte{}
	for _, format := range formats {
		created[format] = markdownToBytes(markdown, format)
	}

	fmt.Println("step 3: open each and check the text")
	for _, format := range formats {
		text := textOf(created[format], format)
		fmt.Println("  format:", format)
		fmt.Println("  chars:", len(text))
		has(text, "North")
	}

	fmt.Println("OK 02_create_from_markdown")
}
