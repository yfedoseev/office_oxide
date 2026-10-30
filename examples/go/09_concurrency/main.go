//go:build office_oxide_dev

// 09_concurrency: open and read many documents from several goroutines at once.
// Run: (cd examples/go/09_concurrency && go run -tags office_oxide_dev .)
package main

import (
	"bytes"
	"fmt"
	"log"
	"os"
	"strings"
	"sync"

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

func reportMarkdown(title string) string {
	return "# " + title + "\n\nRevenue grew in every region.\n\n" +
		"| Region | Q1  | Q2  |\n|--------|-----|-----|\n" +
		"| North  | 120 | 135 |\n| South  | 95  | 110 |\n"
}

// buildDocxBytes returns a DOCX. The Go binding has no DOCX writer: a DOCX
// comes from Markdown, which the library writes to a path, so this uses a
// private temp file for that hop and removes it before returning.
func buildDocxBytes(title string) []byte {
	f := must(os.CreateTemp("", "oo_example_*.docx"))
	path := f.Name()
	f.Close()
	defer os.Remove(path)
	check(oo.CreateFromMarkdown(reportMarkdown(title), "docx", path) == nil, "create the DOCX from Markdown")
	return must(os.ReadFile(path))
}

func main() {
	const workers, documents = 8, 32
	// The host's logger is on for the whole run, writing to a buffer.
	var logs bytes.Buffer
	log.SetOutput(&logs)

	fmt.Println("step 1: build 32 DOCX files in memory")
	docs := make([][]byte, documents)
	for i := range docs {
		docs[i] = buildDocxBytes(fmt.Sprintf("Quarterly Report %d", i))
	}

	fmt.Println("step 2: open and read them from 8 workers")
	var wg sync.WaitGroup
	var mu sync.Mutex
	failures := []string{}
	for w := 0; w < workers; w++ {
		wg.Add(1)
		go func(w int) {
			defer wg.Done()
			for i := w; i < documents; i += workers {
				doc, err := oo.OpenFromBytes(docs[i], "docx")
				if err == nil {
					var text string
					text, err = doc.PlainText()
					doc.Close()
					want := fmt.Sprintf("Quarterly Report %d", i)
					if err == nil && !strings.Contains(text, want) {
						err = fmt.Errorf("document %d: text lacks %q", i, want)
					}
				}
				log.Printf("worker %d read document %d", w, i)
				if err != nil {
					mu.Lock()
					failures = append(failures, err.Error())
					mu.Unlock()
				}
			}
		}(w)
	}
	wg.Wait()
	fmt.Println("  workers:", workers)
	fmt.Println("  documents:", documents)

	fmt.Println("step 3: check every result")
	check(len(failures) == 0, fmt.Sprint("every worker read its own documents back: ", failures))
	check(strings.Count(logs.String(), "\n") == documents, "every document was logged")

	fmt.Println("OK 09_concurrency")
}
