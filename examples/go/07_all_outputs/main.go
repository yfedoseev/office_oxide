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

func main() {
	fmt.Println("step 1: build a DOCX in memory")
	data := buildDocxBytes("Quarterly Report")

	fmt.Println("step 2: open it from bytes")
	doc := must(oo.OpenFromBytes(data, "docx"))
	defer doc.Close()

	fmt.Println("step 3: render every output")
	text := must(doc.PlainText()) // words only, no markup
	md := must(doc.ToMarkdown())  // headings and tables survive as syntax
	html := must(doc.ToHTML())    // a fragment (no <html> wrapper) to embed in a page
	ir := must(doc.ToIRJSON())    // the format-neutral tree, for code that wants structure
	fmt.Printf("  chars: text %d, markdown %d, html %d, ir %d\n", len(text), len(md), len(html), len(ir))

	fmt.Println("step 4: check each output")
	check(len(text) > 0 && len(md) > 0 && len(html) > 0 && len(ir) > 0, "every output is non-empty")
	check(strings.Contains(text, "Quarterly Report"), "plain text has the title")
	check(strings.Contains(md, "# Quarterly Report"), "markdown has the heading")
	check(strings.Contains(html, "<"), "html has markup")
	check(strings.Contains(html, "Quarterly Report"), "html has the title")
	var tree map[string]any
	check(json.Unmarshal([]byte(ir), &tree) == nil, "IR JSON is an object")
	check(strings.Contains(ir, "Quarterly Report"), "IR JSON has the title")

	fmt.Println("OK 07_all_outputs")
}
