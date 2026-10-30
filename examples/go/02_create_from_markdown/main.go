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

// textOf opens bytes as the given format and returns the plain text.
func textOf(data []byte, format string) string {
	doc := must(oo.OpenFromBytes(data, format))
	defer doc.Close()
	return must(doc.PlainText())
}

func main() {
	fmt.Println("step 1: write the Markdown source")
	md := "# Quarterly Report\n\nRevenue grew in every region.\n\n" +
		"| Region | Q1  | Q2  |\n|--------|-----|-----|\n" +
		"| North  | 120 | 135 |\n| South  | 95  | 110 |\n"

	fmt.Println("step 2: create a DOCX, an XLSX and a PPTX from it")
	formats := []string{"docx", "xlsx", "pptx"}
	files := map[string][]byte{}
	for _, format := range formats {
		f := must(os.CreateTemp("", "oo_example_*."+format))
		path := f.Name()
		f.Close()
		defer os.Remove(path)
		check(oo.CreateFromMarkdown(md, format, path) == nil, "create the "+format+" from Markdown")
		files[format] = must(os.ReadFile(path))
	}

	fmt.Println("step 3: open each and check the text")
	for _, format := range formats {
		text := textOf(files[format], format)
		fmt.Println("  format:", format)
		fmt.Println("  chars:", len(text))
		check(strings.Contains(text, "North"), "the table survived")
	}

	fmt.Println("OK 02_create_from_markdown")
}
