//go:build office_oxide_dev

// 03_edit — Edit a document by replacing placeholder text.
//
// Creates a DOCX from Markdown with {{NAME}} and {{DATE}} placeholders,
// opens with oo.OpenEditable, replaces both, saves to a second temp file,
// then reads back and verifies. Exit 0 on success.
//
// Build and run (inside the monorepo, after `cargo build --release --lib`):
//
//	go run -tags office_oxide_dev ./examples/go/03_edit/
package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	oo "github.com/yfedoseev/office_oxide/go"
)

const markdown = `# Invoice

Dear {{NAME}},

Please find attached your invoice for services rendered on {{DATE}}.

## Summary

- Service: Office Oxide Pro License
- Amount: $499.00
- Due date: 30 days from {{DATE}}

Thank you for your business!
`

func main() {
	// Create temp files
	tpl, err := os.CreateTemp("", "oo_03_template_*.docx")
	if err != nil {
		log.Fatal(err)
	}
	tplPath := tpl.Name()
	tpl.Close()
	defer os.Remove(tplPath)

	out, err := os.CreateTemp("", "oo_03_output_*.docx")
	if err != nil {
		log.Fatal(err)
	}
	outPath := out.Name()
	out.Close()
	defer os.Remove(outPath)

	// Create template
	if err := oo.CreateFromMarkdown(markdown, "docx", tplPath); err != nil {
		log.Fatalf("CreateFromMarkdown: %v", err)
	}

	// Edit
	ed, err := oo.OpenEditable(tplPath)
	if err != nil {
		log.Fatalf("OpenEditable: %v", err)
	}
	n1, err := ed.ReplaceText("{{NAME}}", "Alice Smith")
	if err != nil {
		log.Fatalf("ReplaceText NAME: %v", err)
	}
	n2, err := ed.ReplaceText("{{DATE}}", "2026-04-26")
	if err != nil {
		log.Fatalf("ReplaceText DATE: %v", err)
	}
	if err := ed.Save(outPath); err != nil {
		log.Fatalf("Save: %v", err)
	}
	ed.Close()

	fmt.Printf("Replacements: {{NAME}} x%d, {{DATE}} x%d\n", n1, n2)

	// Verify
	doc, err := oo.Open(outPath)
	if err != nil {
		log.Fatalf("Open output: %v", err)
	}
	defer doc.Close()

	text, err := doc.PlainText()
	if err != nil {
		log.Fatalf("PlainText: %v", err)
	}

	if !strings.Contains(text, "Alice Smith") {
		log.Fatal("name replacement failed: 'Alice Smith' not found")
	}
	if !strings.Contains(text, "2026-04-26") {
		log.Fatal("date replacement failed: '2026-04-26' not found")
	}
	if strings.Contains(text, "{{NAME}}") {
		log.Fatal("placeholder {{NAME}} still present")
	}
	if strings.Contains(text, "{{DATE}}") {
		log.Fatal("placeholder {{DATE}} still present")
	}

	fmt.Println("Edit verified successfully.")
	fmt.Println("--- final text ---")
	fmt.Println(text)
}
