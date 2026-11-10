//go:build office_oxide_dev

// 09_concurrency: open and read 32 documents from 8 workers at once, with logging on.
// Run: (cd examples/go/09_concurrency && go run -tags office_oxide_dev .)
package main

import (
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	oo "github.com/yfedoseev/office_oxide/go"
)

const markdown = `# Quarterly Report

Revenue grew in every region.

| Region | Q1  | Q2  |
|--------|-----|-----|
| North  | 120 | 135 |
| South  | 95  | 110 |
`

const workers, documents = 8, 32

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

func main() {
	fmt.Println("step 1: build 32 DOCX files in memory")
	docs := make([][]byte, documents)
	for i := range docs {
		docs[i] = buildDocxBytes(fmt.Sprintf("Quarterly Report %d", i))
	}
	fmt.Println("  documents:", documents)

	fmt.Println("step 2: switch logging on")
	log.SetFlags(log.Lmicroseconds) // the standard logger writes to stderr

	fmt.Println("step 3: parse them on 8 workers")
	type result struct {
		i    int
		text string
		err  error
	}
	results := make(chan result, documents)
	for w := 0; w < workers; w++ {
		go func(w int) {
			for i := w; i < documents; i += workers {
				log.Printf("document %d: parse start", i)
				r := result{i: i}
				doc, err := oo.OpenFromBytes(docs[i], "docx")
				if err == nil {
					r.text, err = doc.PlainText()
					doc.Close()
				}
				r.err = err
				log.Printf("document %d: parse done", i)
				results <- r
			}
		}(w)
	}
	texts := make([]string, documents)
	for range docs {
		select {
		case r := <-results:
			must0(r.err)
			texts[r.i] = r.text
		case <-time.After(60 * time.Second): // a deadlock fails the example instead of hanging it
			check(false, "a parse did not finish: possible deadlock")
		}
	}
	fmt.Println("  workers:", workers)

	fmt.Println("step 4: check every result")
	for i, text := range texts {
		check(text != "", "not every document was read")
		check(strings.Contains(text, fmt.Sprintf("Quarterly Report %d", i)), fmt.Sprintf("document %d came back with the wrong text", i))
	}

	fmt.Println("OK 09_concurrency")
}
