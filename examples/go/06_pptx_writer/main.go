//go:build office_oxide_dev

// 06_pptx_writer: write a presentation with a size, slides, titles, text and an image.
// Run: (cd examples/go/06_pptx_writer && go run -tags office_oxide_dev .)
package main

import (
	"encoding/base64"
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

const png1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="

// buildPptxBytes returns two slides: "Welcome" and "Results" (with an image).
func buildPptxBytes() []byte {
	w := oo.NewPptxWriter()
	defer w.Close()
	w.SetPresentationSize(9144000, 6858000) // 4:3; 914400 EMU = 1 inch
	s1 := w.AddSlide()
	check(w.SetSlideTitle(s1, "Welcome") == nil, "set the slide 1 title")
	check(w.AddSlideText(s1, "Quarterly Report") == nil, "add the slide 1 text")
	s2 := w.AddSlide()
	check(w.SetSlideTitle(s2, "Results") == nil, "set the slide 2 title")
	check(w.AddSlideText(s2, "Revenue grew in every region.") == nil, "add the slide 2 text")
	png := must(base64.StdEncoding.DecodeString(png1x1))
	check(w.AddSlideImage(s2, png, "png", 914400, 1371600, 914400, 914400) == nil, "add the image")
	return must(w.ToBytes())
}

func main() {
	fmt.Println("step 1: build a PPTX in memory")
	data := buildPptxBytes() // see above for the slide-by-slide calls

	fmt.Println("step 2: check that a write to a missing slide is refused")
	w := oo.NewPptxWriter()
	defer w.Close()
	check(w.SetSlideTitle(9, "ghost") != nil, "a missing slide is an error, not a silent drop")

	fmt.Println("step 3: open it from bytes")
	doc := must(oo.OpenFromBytes(data, "pptx"))
	defer doc.Close()
	text := must(doc.PlainText())
	// Plain Markdown drops images; the WithImages variant keeps them inline.
	md := must(doc.ToMarkdownWithImages())
	format := must(doc.Format())
	fmt.Println("  format:", format)
	fmt.Println("  chars:", len(text))
	fmt.Println("  slides: 2")

	fmt.Println("step 4: check the slides")
	check(strings.Contains(text, "Welcome"), "slide 1 title is present")
	check(strings.Contains(text, "Quarterly Report"), "slide 1 text is present")
	check(strings.Contains(text, "Results"), "slide 2 title is present")
	check(strings.Contains(text, "Revenue grew in every region."), "slide 2 text is present")
	check(strings.Contains(md, "[image-base64:"), "the image is present")

	fmt.Println("OK 06_pptx_writer")
}
