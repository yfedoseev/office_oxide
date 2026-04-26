//go:build office_oxide_dev

// 02_create_from_markdown — Convert Markdown to DOCX, XLSX, and PPTX.
//
// For each format: calls oo.CreateFromMarkdown, opens with oo.Open,
// prints a summary. Exit 0 on success.
//
// Build and run (inside the monorepo, after `cargo build --release --lib`):
//
//	go run -tags office_oxide_dev ./examples/go/02_create_from_markdown/
package main

import (
	"fmt"
	"log"
	"os"

	oo "github.com/yfedoseev/office_oxide/go"
)

const markdown = `# Quarterly Report

Generated automatically from Markdown using Office Oxide.

## Highlights

- Revenue grew by **32%** year-over-year
- Customer satisfaction: 4.8 / 5.0
- New products launched: Widget Pro, Widget Lite

## Financial Summary

| Category   | Q3 2025 | Q4 2025 |
|------------|---------|---------|
| Revenue    | $1.2M   | $1.6M   |
| Expenses   | $0.8M   | $0.9M   |
| Net Profit | $0.4M   | $0.7M   |
`

var formats = []string{"docx", "xlsx", "pptx"}

func main() {
	tmps := make(map[string]string)

	// Create temp files
	for _, ext := range formats {
		f, err := os.CreateTemp("", "oo_02_create_*."+ext)
		if err != nil {
			log.Fatalf("TempFile: %v", err)
		}
		tmps[ext] = f.Name()
		f.Close()
	}
	defer func() {
		for _, p := range tmps {
			os.Remove(p)
		}
	}()

	for _, format := range formats {
		path := tmps[format]
		if err := oo.CreateFromMarkdown(markdown, format, path); err != nil {
			log.Fatalf("CreateFromMarkdown(%s): %v", format, err)
		}

		doc, err := oo.Open(path)
		if err != nil {
			log.Fatalf("Open(%s): %v", format, err)
		}

		text, err := doc.PlainText()
		doc.Close()
		if err != nil {
			log.Fatalf("PlainText(%s): %v", format, err)
		}

		if len(text) == 0 {
			log.Fatalf("%s: plain text is empty", format)
		}

		preview := text
		if len(preview) > 100 {
			preview = preview[:100]
		}

		fmt.Printf("\n=== %s ===\n", format)
		fmt.Printf("plain text length: %d chars\n", len(text))
		fmt.Printf("first 100 chars: %q\n", preview)
	}

	fmt.Println("\nAll formats created and verified.")
}
