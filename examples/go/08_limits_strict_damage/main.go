//go:build office_oxide_dev

// 08_limits_strict_damage: strict mode refuses a damaged archive; a lenient open reports the damage.
// Run: (cd examples/go/08_limits_strict_damage && go run -tags office_oxide_dev .)
package main

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"os"
	"slices"
	"strings"

	oo "github.com/yfedoseev/office_oxide/go"
)

const markdown = `# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`

const part = "word/_rels/document.xml.rels"

// damagePart overwrites the start of one part's compressed data. A first
// deflate byte of 0xFF declares a reserved block type, so every inflater
// rejects the part, and the central directory stays intact. The part is found
// by name, not by offset, so the recipe does not depend on how big the other
// parts are.
func damagePart(data []byte) []byte {
	at := bytes.Index(data, []byte(part)) // the first hit is the part's local header
	check(at > 0, "the part to damage is missing from the archive")
	// The extra-field length sits in the two bytes just before the name.
	start := at + len(part) + int(binary.LittleEndian.Uint16(data[at-2:at]))
	damaged := append([]byte(nil), data...)
	for i := start; i < start+16; i++ {
		damaged[i] = 0xFF
	}
	return damaged
}

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

// markdownToBytes creates a document from Markdown. The library writes to a
// path, so the file lives in a private temp file just long enough to read its
// bytes back.
func markdownToBytes(md, format string) []byte {
	f := must(os.CreateTemp("", "oo_example_*."+format))
	path := f.Name()
	f.Close()
	defer os.Remove(path)
	must0(oo.CreateFromMarkdown(md, format, path))
	return must(os.ReadFile(path))
}

func buildDocxBytes(title string) []byte {
	return markdownToBytes(strings.ReplaceAll(markdown, "Quarterly Report", title), "docx")
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

func main() {
	fmt.Println("step 1: build a DOCX in memory")
	good := buildDocxBytes("Quarterly Report")

	fmt.Println("step 2: damage one part: overwrite 16 bytes of " + part)
	damaged := damagePart(good)

	strict := must(oo.NewStrictParseOptions())
	defer strict.Close()
	lenient := must(oo.NewParseOptions())
	defer lenient.Close()

	fmt.Println("step 3: open it in strict mode")
	doc, err := oo.OpenFromBytesWith(damaged, "docx", strict)
	if err == nil {
		doc.Close()
	}
	check(err != nil, "strict mode accepted a damaged part")
	fmt.Println("  error:", err)

	fmt.Println("step 4: open it leniently")
	doc = must(oo.OpenFromBytesWith(damaged, "docx", lenient))
	kinds := []string{}
	for _, d := range must(doc.Damage()) {
		kinds = append(kinds, d.Kind)
	}
	text := must(doc.PlainText())
	doc.Close()
	fmt.Println("  damage:", strings.Join(kinds, ", "))
	check(slices.Contains(kinds, "unreadable_part"), "the damage report does not name an unreadable part")
	has(text, "Quarterly Report")

	fmt.Println("step 5: truncate the archive to its first half and open it both ways")
	half := good[:len(good)/2]
	for _, c := range []struct {
		options *oo.ParseOptions
		message string
	}{
		{strict, "strict mode accepted a truncated archive"},
		{lenient, "a lenient open accepted a truncated archive"},
	} {
		doc, err = oo.OpenFromBytesWith(half, "docx", c.options)
		if err == nil {
			doc.Close()
		}
		check(err != nil, c.message)
		fmt.Println("  error:", err)
	}

	fmt.Println("step 6: a limit applies to one document only")
	limited := must(oo.NewParseOptions())
	defer limited.Close()
	must0(limited.Set("part_bytes", 100))
	doc, err = oo.OpenFromBytesWith(good, "docx", limited)
	if err == nil {
		doc.Close()
	}
	check(err != nil, "the part_bytes limit was not applied")
	fmt.Println("  error:", err)
	has(textOf(good, "docx"), "Quarterly Report") // the default limits still accept it

	fmt.Println("OK 08_limits_strict_damage")
}
