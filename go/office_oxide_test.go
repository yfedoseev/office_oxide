//go:build office_oxide_dev

package officeoxide

import (
	"encoding/json"
	"errors"
	"math"
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

func workbookText(t *testing.T, data []byte) string {
	t.Helper()
	doc, err := OpenFromBytes(data, "xlsx")
	if err != nil {
		t.Fatalf("OpenFromBytes: %v", err)
	}
	defer doc.Close()
	txt, err := doc.PlainText()
	if err != nil {
		t.Fatalf("PlainText: %v", err)
	}
	return txt
}

// The writer status codes were discarded, so a write to a missing sheet or
// slide, or outside the grid, reported success while writing nothing.
func TestWriterStatusesAreErrors(t *testing.T) {
	w := NewXlsxWriter()
	defer w.Close()
	w.AddSheet("S")
	var oe *Error
	if err := w.SetCell(5, 0, 0, "x"); !errors.As(err, &oe) || oe.Code != 6 {
		t.Fatalf("missing sheet: got %v", err)
	}
	if err := w.SetCell(0, 1_048_576, 0, "x"); err == nil {
		t.Fatal("row outside the grid must fail")
	}
	if err := w.SetCellStyled(5, 0, 0, "x", true, ""); err == nil {
		t.Fatal("styled write to a missing sheet must fail")
	}
	if err := w.SetCell(0, 0, 0, "ok"); err != nil {
		t.Fatalf("valid write: %v", err)
	}

	p := NewPptxWriter()
	defer p.Close()
	p.AddSlide()
	if err := p.SetSlideTitle(99, "x"); err == nil {
		t.Fatal("SetSlideTitle on a missing slide must fail")
	}
	if err := p.AddSlideText(99, "x"); err == nil {
		t.Fatal("AddSlideText on a missing slide must fail")
	}
	if err := p.AddSlideImage(0, []byte{1, 2, 3}, "bmp", 0, 0, 1, 1); err == nil {
		t.Fatal("AddSlideImage with an unknown format must fail")
	}
	if err := p.SetSlideTitle(0, "ok"); err != nil {
		t.Fatalf("valid title: %v", err)
	}
}

// Booleans, formulas and exact integers.
func TestWriterValueKinds(t *testing.T) {
	w := NewXlsxWriter()
	defer w.Close()
	w.AddSheet("S")
	must := func(err error) {
		t.Helper()
		if err != nil {
			t.Fatal(err)
		}
	}
	must(w.SetCell(0, 0, 0, true))
	must(w.SetCell(0, 1, 0, 2))
	must(w.SetFormula(0, 2, 0, "=SUM(A2:A2)"))
	must(w.SetCell(0, 3, 0, Formula("A2*2")))
	must(w.SetCell(0, 4, 0, "=literal"))
	must(w.SetCell(0, 5, 0, int64(1)<<53))
	must(w.SetCell(0, 6, 0, int64(math.MinInt64+1)>>10))
	for _, big := range []any{int64(1)<<53 + 1, uint64(math.MaxUint64), int64(math.MinInt64)} {
		if err := w.SetCell(0, 7, 0, big); err == nil {
			t.Fatalf("%v must be rejected: a double cannot hold it exactly", big)
		}
	}
	data, err := w.ToBytes()
	must(err)
	doc, err := OpenFromBytes(data, "xlsx")
	must(err)
	defer doc.Close()
	ir, err := doc.ToIRJSON()
	must(err)
	for _, want := range []string{"SUM(A2:A2)", "A2*2", "=literal"} {
		if !strings.Contains(ir, want) {
			t.Fatalf("IR lacks %q: %s", want, ir)
		}
	}
	if txt := workbookText(t, data); !strings.Contains(txt, "TRUE") {
		t.Fatalf("boolean not written as TRUE: %q", txt)
	}
}

// In-memory editing needed a temporary file: open_from_bytes existed in
// the FFI but not here.
func TestEditableFromBytes(t *testing.T) {
	w := NewXlsxWriter()
	defer w.Close()
	w.AddSheet("S")
	if err := w.SetCell(0, 0, 0, "old"); err != nil {
		t.Fatal(err)
	}
	data, err := w.ToBytes()
	if err != nil {
		t.Fatal(err)
	}
	ed, err := OpenEditableFromBytes(data, "xlsx")
	if err != nil {
		t.Fatalf("OpenEditableFromBytes: %v", err)
	}
	defer ed.Close()
	if err := ed.SetCell(0, "A1", NewStringCell("new")); err != nil {
		t.Fatal(err)
	}
	var oe *Error
	if _, err := ed.ReplaceText("", "x"); !errors.As(err, &oe) || oe.Code != 1 {
		t.Fatalf("empty find: got %v, want invalid argument", err)
	}
	out, err := ed.SaveToBytes()
	if err != nil {
		t.Fatal(err)
	}
	if txt := workbookText(t, out); !strings.Contains(txt, "new") {
		t.Fatalf("edit lost: %q", txt)
	}
}

func TestToMarkdownWithImages(t *testing.T) {
	png := []byte{
		0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
		0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
		0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xdd, 0x8d, 0xb0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
		0x44, 0xae, 0x42, 0x60, 0x82,
	}
	p := NewPptxWriter()
	defer p.Close()
	p.AddSlide()
	if err := p.AddSlideImage(0, png, "png", 0, 0, 914400, 914400); err != nil {
		t.Fatal(err)
	}
	data, err := p.ToBytes()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := OpenFromBytes(data, "pptx")
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	md, err := doc.ToMarkdownWithImages()
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(md, "[image-base64:") {
		t.Fatalf("no embedded image: %q", md)
	}
}
