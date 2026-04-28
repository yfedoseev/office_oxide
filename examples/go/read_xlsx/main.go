package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"

	oo "github.com/yfedoseev/office_oxide/go"
)

type ir struct {
	Sections []struct {
		Title    *string `json:"title"`
		Elements []struct {
			Type string `json:"type"`
			Rows []struct {
				Cells []struct {
					Text string `json:"text"`
				} `json:"cells"`
			} `json:"rows"`
		} `json:"elements"`
	} `json:"sections"`
}

func main() {
	if len(os.Args) != 2 {
		log.Fatal("usage: read_xlsx <file.xlsx>")
	}
	doc, err := oo.Open(os.Args[1])
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Close()

	raw, err := doc.ToIRJSON()
	if err != nil {
		log.Fatal(err)
	}
	var parsed ir
	if err := json.Unmarshal([]byte(raw), &parsed); err != nil {
		log.Fatal(err)
	}
	for i, s := range parsed.Sections {
		title := ""
		if s.Title != nil {
			title = *s.Title
		}
		fmt.Printf("# sheet %d: %s\n", i, title)
		for _, el := range s.Elements {
			if el.Type != "table" {
				continue
			}
			for _, row := range el.Rows {
				for j, c := range row.Cells {
					if j > 0 {
						fmt.Print("\t")
					}
					fmt.Print(c.Text)
				}
				fmt.Println()
			}
		}
	}
}
