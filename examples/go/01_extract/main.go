//go:build office_oxide_dev

// 01_extract — Self-contained extract demo using the Go bindings.
//
// Creates a DOCX from Markdown (written to a temp file), opens it with
// oo.Open, and prints the format, plain text, and Markdown output.
// Exit 0 on success.
//
// Build and run (inside the monorepo, after `cargo build --release --lib`):
//
//	go run -tags office_oxide_dev ./examples/go/01_extract/
package main

import (
	"fmt"
	"log"
	"os"

	oo "github.com/yfedoseev/office_oxide/go"
)

const markdown = `# Office Oxide Extract Demo

This document was created from Markdown and parsed back via Go CGo bindings.

## Features

- Plain text extraction
- Markdown conversion
- IR (intermediate representation) access
`

func main() {
	// Create temp file
	f, err := os.CreateTemp("", "oo_01_extract_*.docx")
	if err != nil {
		log.Fatal(err)
	}
	tmpPath := f.Name()
	f.Close()
	defer os.Remove(tmpPath)

	// Create DOCX from Markdown
	if err := oo.CreateFromMarkdown(markdown, "docx", tmpPath); err != nil {
		log.Fatalf("CreateFromMarkdown: %v", err)
	}

	// Open and extract
	doc, err := oo.Open(tmpPath)
	if err != nil {
		log.Fatalf("Open: %v", err)
	}
	defer doc.Close()

	format, err := doc.Format()
	if err != nil {
		log.Fatalf("Format: %v", err)
	}

	text, err := doc.PlainText()
	if err != nil {
		log.Fatalf("PlainText: %v", err)
	}

	md, err := doc.ToMarkdown()
	if err != nil {
		log.Fatalf("ToMarkdown: %v", err)
	}

	irJSON, err := doc.ToIRJSON()
	if err != nil {
		log.Fatalf("ToIRJSON: %v", err)
	}

	fmt.Println("format:", format)
	fmt.Println("--- plain text (first 200 chars) ---")
	preview := text
	if len(preview) > 200 {
		preview = preview[:200]
	}
	fmt.Println(preview)
	fmt.Println("--- markdown length:", len(md), "chars ---")
	fmt.Println("--- IR JSON length:", len(irJSON), "chars ---")

	// Verify
	if format != "docx" {
		log.Fatalf("expected format=docx, got %q", format)
	}
	if len(text) == 0 {
		log.Fatal("plain text is empty")
	}

	fmt.Println("\nAll checks passed.")
}
