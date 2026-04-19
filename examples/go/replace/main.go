package main

import (
	"fmt"
	"log"
	"os"

	oo "github.com/yfedoseev/office_oxide/go"
)

func main() {
	if len(os.Args) != 3 {
		log.Fatal("usage: replace <template> <output>")
	}
	ed, err := oo.OpenEditable(os.Args[1])
	if err != nil {
		log.Fatal(err)
	}
	defer ed.Close()

	n, err := ed.ReplaceText("{{NAME}}", "Alice")
	if err != nil {
		log.Fatal(err)
	}
	m, _ := ed.ReplaceText("{{DATE}}", "2026-04-18")
	fmt.Println("replacements:", n+m)
	if err := ed.Save(os.Args[2]); err != nil {
		log.Fatal(err)
	}
}
