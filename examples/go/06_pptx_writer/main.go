//go:build office_oxide_dev

// 06_pptx_writer: write a presentation with a size, slides, titles, text and an image, then read it back.
// Run: (cd examples/go/06_pptx_writer && go run -tags office_oxide_dev .)
package main

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"os"
	"strings"

	oo "github.com/yfedoseev/office_oxide/go"
)

const png1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="

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

func openBytes(data []byte, format string) *oo.Document {
	return must(oo.OpenFromBytes(data, format))
}

// textOf opens bytes as the given format and returns the plain text.
func textOf(data []byte, format string) string {
	doc := openBytes(data, format)
	defer doc.Close()
	return must(doc.PlainText())
}

func buildPptxBytes() []byte {
	w := oo.NewPptxWriter()
	defer w.Close()
	w.SetPresentationSize(12192000, 6858000) // 16:9, in EMU (914400 per inch)
	one := w.AddSlide()
	must0(w.SetSlideTitle(one, "Welcome"))
	must0(w.AddSlideText(one, "Quarterly Report"))
	two := w.AddSlide()
	must0(w.SetSlideTitle(two, "Results"))
	must0(w.AddSlideText(two, "Revenue grew in every region."))
	png := must(base64.StdEncoding.DecodeString(png1x1))
	must0(w.AddSlideImage(two, png, "png", 914400, 2743200, 1828800, 1828800))
	return must(w.ToBytes())
}

// sectionCount counts the sheets or slides in an IR JSON tree.
func sectionCount(irJSON string) int {
	var tree struct{ Sections []json.RawMessage }
	must0(json.Unmarshal([]byte(irJSON), &tree))
	return len(tree.Sections)
}

func main() {
	fmt.Println("step 1: build a PPTX in memory")
	data := buildPptxBytes()

	fmt.Println("step 2: open it from bytes")
	doc := openBytes(data, "pptx")
	defer doc.Close()
	text := must(doc.PlainText())
	slides := sectionCount(must(doc.ToIRJSON()))
	// Plain Markdown drops images; the WithImages variant keeps them in place.
	withImages := must(doc.ToMarkdownWithImages())
	fmt.Println("  format:", must(doc.Format()))
	fmt.Println("  chars:", len(text))
	fmt.Println("  slides:", slides)

	fmt.Println("step 3: check the slides")
	check(slides == 2, "expected two slides")
	for _, expected := range []string{"Welcome", "Quarterly Report", "Results", "Revenue grew in every region."} {
		has(text, expected)
	}
	check(strings.Index(text, "Welcome") < strings.Index(text, "Results"), "the slides are out of order")
	check(strings.Contains(withImages, "[image-base64:"), "the slide image is missing")

	fmt.Println("OK 06_pptx_writer")
}
