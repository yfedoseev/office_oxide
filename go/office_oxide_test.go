//go:build office_oxide_dev

package officeoxide

import (
	"encoding/json"
	"os"
	"strings"
	"testing"
)

// Tests require a pre-built fixture at /tmp/ffi_smoke.docx. See the monorepo
// smoke instructions or build one with the Rust `create` API.

func TestVersion(t *testing.T) {
	v := Version()
	if v == "" {
		t.Fatal("expected non-empty version")
	}
	t.Logf("office_oxide version: %s", v)
}

func TestDetectFormat(t *testing.T) {
	if DetectFormat("x.docx") != "docx" {
		t.Fatal("expected docx")
	}
	if DetectFormat("x.unknown") != "" {
		t.Fatal("expected empty for unknown ext")
	}
}

func TestOpenAndExtract(t *testing.T) {
	fixture := "/tmp/ffi_smoke.docx"
	if _, err := os.Stat(fixture); err != nil {
		t.Skipf("fixture %s missing: %v", fixture, err)
	}
	doc, err := Open(fixture)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer doc.Close()

	f, err := doc.Format()
	if err != nil || f != "docx" {
		t.Fatalf("Format: got %q err=%v", f, err)
	}
	text, err := doc.PlainText()
	if err != nil {
		t.Fatalf("PlainText: %v", err)
	}
	if !strings.Contains(text, "Hello") {
		t.Fatalf("unexpected text: %q", text)
	}
	md, err := doc.ToMarkdown()
	if err != nil {
		t.Fatalf("ToMarkdown: %v", err)
	}
	if !strings.Contains(md, "# ") {
		t.Fatalf("expected markdown heading, got: %q", md)
	}
	irJSON, err := doc.ToIRJSON()
	if err != nil {
		t.Fatalf("ToIRJSON: %v", err)
	}
	var ir map[string]any
	if err := json.Unmarshal([]byte(irJSON), &ir); err != nil {
		t.Fatalf("IR JSON parse: %v", err)
	}
	if _, ok := ir["sections"]; !ok {
		t.Fatalf("missing sections in IR: %v", ir)
	}
}

func TestEditableReplaceText(t *testing.T) {
	fixture := "/tmp/ffi_smoke.docx"
	if _, err := os.Stat(fixture); err != nil {
		t.Skipf("fixture %s missing: %v", fixture, err)
	}
	ed, err := OpenEditable(fixture)
	if err != nil {
		t.Fatalf("OpenEditable: %v", err)
	}
	defer ed.Close()
	n, err := ed.ReplaceText("Hello", "Greetings")
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}
	if n < 1 {
		t.Fatalf("expected at least 1 replacement, got %d", n)
	}
	out := "/tmp/ffi_smoke_go_edit.docx"
	if err := ed.Save(out); err != nil {
		t.Fatalf("Save: %v", err)
	}
	txt, err := ExtractText(out)
	if err != nil {
		t.Fatalf("ExtractText: %v", err)
	}
	if !strings.Contains(txt, "Greetings") {
		t.Fatalf("replacement not persisted: %q", txt)
	}
}
