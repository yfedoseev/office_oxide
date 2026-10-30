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
	format := must(doc.Format())
	fmt.Println("  format:", format)
	check(format == "docx", "format is docx")

	fmt.Println("step 3: extract plain text, Markdown and IR JSON")
	text := must(doc.PlainText())
	md := must(doc.ToMarkdown())
	ir := must(doc.ToIRJSON())
	fmt.Println("  chars:", len(text))

	fmt.Println("step 4: check the output")
	check(strings.Contains(text, "Revenue grew in every region."), "plain text has the paragraph")
	check(strings.Contains(md, "Quarterly Report"), "markdown has the title")
	check(strings.HasPrefix(ir, "{"), "IR JSON is an object")

	fmt.Println("OK 01_extract")
}
