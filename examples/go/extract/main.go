package main

import (
	"fmt"
	"log"
	"os"

	oo "github.com/yfedoseev/office_oxide/go"
)

func main() {
	if len(os.Args) != 2 {
		log.Fatal("usage: extract <file>")
	}
	doc, err := oo.Open(os.Args[1])
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()

	format, _ := doc.Format()
	fmt.Println("format:", format)
	text, err := doc.PlainText()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("--- plain text ---")
	fmt.Println(text)
	md, _ := doc.ToMarkdown()
	fmt.Println("--- markdown ---")
	fmt.Println(md)
}
