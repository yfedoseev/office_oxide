//go:build office_oxide_dev

// 08_limits_strict_damage: strict options refuse a truncated archive; a lenient open recovers or fails cleanly.
// Run: (cd examples/go/08_limits_strict_damage && go run -tags office_oxide_dev .)
package main

import (
	"fmt"
	"os"

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
	fmt.Println("step 1: build a DOCX in memory")
	good := buildDocxBytes("Quarterly Report")

	fmt.Println("step 2: keep only the first half of its bytes")
	cut := good[:len(good)/2]

	fmt.Println("step 3: open the damaged bytes with strict options")
	strict := must(oo.NewStrictParseOptions())
	defer strict.Close()
	lenient := must(oo.NewParseOptions())
	defer lenient.Close()
	check(strict.Set("text_chars", 1000000) == nil, "a known option is accepted")
	check(strict.Set("no_such_option", 1) != nil, "an unknown option is refused")
	doc, err := oo.OpenFromBytesWith(cut, "docx", strict)
	check(err != nil, "strict mode refuses a truncated archive")
	fmt.Println("  error:", err)

	fmt.Println("step 4: open the damaged bytes leniently")
	doc, err = oo.OpenFromBytesWith(cut, "docx", lenient)
	if err == nil {
		damage := must(doc.Damage())
		doc.Close()
		fmt.Println("  damage:", damage)
		check(len(damage) > 0, "a half archive is not reported as sound")
	} else {
		fmt.Println("  error:", err)
	}

	fmt.Println("step 5: open the intact bytes with strict options")
	doc = must(oo.OpenFromBytesWith(good, "docx", strict))
	defer doc.Close()
	damage := must(doc.Damage())
	fmt.Println("  damage:", damage)
	check(len(damage) == 0, "an intact file reports no damage")

	fmt.Println("OK 08_limits_strict_damage")
}
