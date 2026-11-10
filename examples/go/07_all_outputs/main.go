//go:build office_oxide_dev

// 07_all_outputs: render one document as plain text, Markdown, HTML and IR JSON.
// Run: (cd examples/go/07_all_outputs && go run -tags office_oxide_dev .)
package main

import (
	"encoding/json"
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
	fmt.Println("  format:", must(doc.Format()))

	fmt.Println("step 3: render every output")
	outputs := []struct{ name, value string }{
		{"plain", must(doc.PlainText())},     // words only: no markup, no structure
		{"markdown", must(doc.ToMarkdown())}, // headings and pipe tables survive as syntax
		{"html", must(doc.ToHTML())},         // a fragment (no <html> wrapper), tables as <table>
		{"ir_json", must(doc.ToIRJSON())},    // the format-neutral tree as JSON, for programs
	}
	for _, o := range outputs {
		fmt.Printf("  chars: %s %d\n", o.name, len(o.value))
	}

	fmt.Println("step 4: check each output")
	for _, o := range outputs {
		check(strings.TrimSpace(o.value) != "", "the "+o.name+" output is empty")
	}
	has(outputs[0].value, "Quarterly Report")
	check(strings.Contains(outputs[1].value, "# Quarterly Report"), "the Markdown is missing the heading")
	check(strings.Contains(outputs[2].value, "<table"), "the HTML has no table")
	var tree map[string]any
	check(json.Unmarshal([]byte(outputs[3].value), &tree) == nil, "the IR JSON does not parse to an object")

	fmt.Println("OK 07_all_outputs")
}
