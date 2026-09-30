//! Text and formatting fidelity of the legacy `.doc` reader, on synthetic
//! documents built in code by `tests/common`.

mod common;

use common::{FibTweaks, Para, Subdocs, build_doc_full, open_doc, prose_grpprl};
use office_oxide::ir::{Element, InlineContent};

fn para(text: &'static str) -> Para {
    Para {
        text,
        terminator: '\r',
        grpprl: prose_grpprl(),
    }
}

fn all_text(ir: &office_oxide::ir::DocumentIR) -> String {
    fn walk(e: &Element, out: &mut String) {
        match e {
            Element::Paragraph(p) => {
                for c in &p.content {
                    if let InlineContent::Text(t) = c {
                        out.push_str(&t.text);
                    }
                }
                out.push('\n');
            },
            Element::Heading(h) => {
                for c in &h.content {
                    if let InlineContent::Text(t) = c {
                        out.push_str(&t.text);
                    }
                }
                out.push('\n');
            },
            Element::Footnote(n) | Element::Endnote(n) => {
                for e in &n.content {
                    walk(e, out);
                }
            },
            _ => {},
        }
    }
    let mut out = String::new();
    for s in &ir.sections {
        for e in &s.elements {
            walk(e, &mut out);
        }
    }
    out
}

/// Main-story reference marks — `0x02` (auto-numbered footnote/endnote
/// reference) and `0x05` (comment reference) — are anchors, not text, and
/// leaked into `plain_text()` and the IR as literal U+0002/U+0005.
/// `0x1E` is a non-breaking hyphen (drawn, U+2011) and `0x1F` an optional
/// hyphen (a line-break hint, dropped as the DOCX reader drops
/// `w:softHyphen`).
#[test]
fn test_reference_marks_and_hyphen_controls_do_not_leak_into_text() {
    let paras = [para(
        "See note\u{2} and comment\u{5}; well\u{1E}known co\u{1F}operation.",
    )];
    let subdocs = Subdocs {
        footnotes: "\u{2} First note.\r\u{2} Second note.",
        ..Default::default()
    };
    let doc = open_doc(&build_doc_full(&paras, &subdocs, FibTweaks::default()));
    let text = doc.plain_text();
    for bad in ['\u{2}', '\u{5}', '\u{1E}', '\u{1F}'] {
        assert!(!text.contains(bad), "U+{:04X} leaked: {text:?}", bad as u32);
    }
    assert!(
        text.contains("See note and comment; well\u{2011}known cooperation."),
        "{text:?}"
    );

    let ir = doc.to_ir();
    let ir_text = all_text(&ir);
    for bad in ['\u{2}', '\u{5}', '\u{1E}', '\u{1F}'] {
        assert!(!ir_text.contains(bad), "U+{:04X} leaked into IR: {ir_text:?}", bad as u32);
    }
    // The footnote bodies are still split one per note.
    let notes: Vec<_> = ir.sections[0]
        .elements
        .iter()
        .filter(|e| matches!(e, Element::Footnote(_)))
        .collect();
    assert_eq!(notes.len(), 2, "{:?}", ir.sections[0].elements);
}

/// A nested table (`sprmPItap` > 1) is flattened into its outer table.
/// That used to be announced by a paragraph of text the document does not
/// have; it is a metadata warning now, and the document's own text is
/// untouched.
#[test]
fn test_flattened_nested_table_is_a_warning_not_document_text() {
    use common::{cell_grpprl, row_grpprl};
    // Depth-2 cell and row mark: sprmPItap operand = 2.
    let mut cell = cell_grpprl();
    cell[5] = 2;
    let mut row = row_grpprl(&[0, 1000], &[0]);
    row[8] = 2;
    let paras = [
        Para {
            text: "Inner",
            terminator: '\u{7}',
            grpprl: cell,
        },
        Para {
            text: "",
            terminator: '\u{7}',
            grpprl: row,
        },
        para("After."),
    ];
    let doc = open_doc(&build_doc_full(&paras, &Subdocs::default(), FibTweaks::default()));
    let ir = doc.to_ir();
    let md = ir.to_markdown();
    assert!(!md.contains("nested table"), "fabricated text in the output: {md}");
    assert!(md.contains("Inner"), "{md}");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("nested table")),
        "{:?}",
        ir.metadata.warnings
    );
}

fn header_text(hf: &Option<office_oxide::ir::HeaderFooter>) -> String {
    let mut out = String::new();
    for e in hf.iter().flat_map(|h| h.content.iter()) {
        if let Element::Paragraph(p) = e {
            for c in &p.content {
                if let InlineContent::Text(t) = c {
                    out.push_str(&t.text);
                }
            }
        }
    }
    out
}

/// A three-section document: every section's headers are read from its
/// own `PlcfHdd` story group ([MS-DOC] `Plcfhdd`), an
/// empty story inherits the previous section's, and each section becomes
/// its own IR `Section`. Only section 1's headers were read, one `Section`
/// was built, and a section's last paragraph (ended by its section mark,
/// 0x0C) merged with the next section's first.
#[test]
fn test_every_section_gets_its_own_headers_and_ir_section() {
    let paras = [
        Para {
            text: "Sec one.",
            terminator: '\u{C}',
            grpprl: prose_grpprl(),
        },
        Para {
            text: "Sec two.",
            terminator: '\u{C}',
            grpprl: prose_grpprl(),
        },
        para("Sec three."),
    ];
    let subdocs = Subdocs {
        headers: "H1\rH2",
        ..Default::default()
    };
    // Header document "H1\rH2\r" plus the builder's final mark: 7 CPs.
    // 6 separator stories, then 6 per section; only the odd headers of
    // sections 1 and 2 have text, section 3's group is all empty.
    let mut hdd = vec![0u32; 7]; // separators + sec1 even header start
    hdd.extend([0, 3, 3, 3, 3, 3]); // sec1 odd hdr = [0,3), rest empty
    hdd.extend([3, 6, 6, 6, 6, 6]); // sec2 odd hdr = [3,6)
    hdd.extend([6, 6, 6, 6, 6, 6]); // sec3: nothing
    hdd.push(6); // trailing aCP (ignored)
    assert_eq!(hdd.len(), 6 + 18 + 2);
    let tweaks = FibTweaks {
        plcf_sed: vec![0, 9, 18, 29],
        plcf_hdd: hdd,
        ..Default::default()
    };
    let doc = open_doc(&build_doc_full(&paras, &subdocs, tweaks));
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 3, "{:#?}", ir.sections);
    assert_eq!(header_text(&ir.sections[0].header), "H1");
    assert_eq!(header_text(&ir.sections[1].header), "H2");
    assert_eq!(header_text(&ir.sections[2].header), "H2", "inherited from section 2");
    let body = |i: usize| {
        all_text(&office_oxide::ir::DocumentIR {
            sections: vec![ir.sections[i].clone()],
            ..Default::default()
        })
    };
    assert!(body(0).contains("Sec one.") && !body(0).contains("Sec two."), "{}", body(0));
    assert!(body(1).contains("Sec two.") && !body(1).contains("Sec three."), "{}", body(1));
    assert!(body(2).contains("Sec three."), "{}", body(2));
}

/// `PlcftxbxTxt` ([MS-DOC] `PlcftxbxTxt`) delimits each text box's
/// story within the text-box subdocument. It was never read, so every
/// box's text came out as one merged blob; each box is its own element now.
#[test]
fn test_text_box_stories_are_split_per_box() {
    let subdocs = Subdocs {
        textboxes: "Box one\rBox two",
        ..Default::default()
    };
    // "Box one\r" [0,8), "Box two\r" [8,16), then the trailing dummy story.
    let tweaks = FibTweaks {
        plcf_txbx_txt: vec![0, 8, 16, 16],
        ..Default::default()
    };
    let doc = open_doc(&build_doc_full(&[para("Body.")], &subdocs, tweaks));
    let ir = doc.to_ir();
    let boxes: Vec<String> = ir.sections[0]
        .elements
        .iter()
        .filter_map(|e| match e {
            Element::TextBox(tb) => Some(all_text(&office_oxide::ir::DocumentIR {
                sections: vec![office_oxide::ir::Section {
                    elements: tb.content.clone(),
                    ..Default::default()
                }],
                ..Default::default()
            })),
            _ => None,
        })
        .collect();
    assert_eq!(boxes, ["Box one\n", "Box two\n"]);
}

/// Word 6.0/95 stores pictures in the `WordDocument` stream as a `PICF`
/// header ([MS-DOC] §2.9.192) followed by the metafile; none were ever
/// extracted, since only the Word 97 `Data` stream was scanned.
#[test]
fn test_word6_picf_metafile_picture_is_extracted() {
    let mut wmf = Vec::new();
    wmf.extend_from_slice(&1u16.to_le_bytes()); // Type = memory
    wmf.extend_from_slice(&9u16.to_le_bytes()); // HeaderSize
    wmf.extend_from_slice(&0x0300u16.to_le_bytes()); // Version
    wmf.extend_from_slice(&12u32.to_le_bytes()); // Size: 24 bytes
    wmf.extend_from_slice(&0u16.to_le_bytes());
    wmf.extend_from_slice(&3u32.to_le_bytes());
    wmf.extend_from_slice(&0u16.to_le_bytes());
    wmf.extend_from_slice(&[3, 0, 0, 0, 0, 0]); // META_EOF
    let mut picf = vec![0u8; 0x44];
    picf[0..4].copy_from_slice(&((0x44 + wmf.len()) as i32).to_le_bytes());
    picf[4..6].copy_from_slice(&0x44u16.to_le_bytes());
    picf[6..8].copy_from_slice(&8u16.to_le_bytes()); // MM_ANISOTROPIC
    picf.extend_from_slice(&wmf);
    let bytes = common::build_word6_doc_with_tail(0xA5DC, b"Picture below.\r", &picf);
    let ir = open_doc(&bytes).to_ir();
    let images: Vec<_> = ir
        .sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .filter_map(|e| match e {
            Element::Image(i) => i.data.clone(),
            _ => None,
        })
        .collect();
    assert_eq!(images, vec![wmf]);
}
