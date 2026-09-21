use std::io::{BufWriter, Write};

use office_oxide::Document;

/// Dump the document IR as JSON — the IR's own serde form, byte-for-byte
/// what the Python, MCP, C and WASM surfaces already emit.
///
/// The IR is serialized straight into a buffered stdout. There used to be a
/// hand-rolled `serde_json::Value` projection here, which cost two things:
/// it had to be kept in step with the IR by hand (five separate "field
/// added, projection missed it" defects, each caught only by a corpus
/// sweep), and it built the whole tree in memory — one owned-key map per
/// cell, paragraph and span — before pretty-printing it into a second,
/// whole-document string. On a 4.7 MB `.xls` with 316k cells that was
/// 2.5 GB of RSS and 18 s; on a 14 MB one, 9.8 GB and an OOM kill. The IR
/// itself for those files is a quarter of a gigabyte.
pub fn run(file: &str) -> Result<(), Box<dyn std::error::Error>> {
    let doc = Document::open(file)?;
    let ir = doc.to_ir();
    let stdout = std::io::stdout();
    let mut out = BufWriter::new(stdout.lock());
    serde_json::to_writer_pretty(&mut out, &ir)?;
    out.write_all(b"\n")?;
    out.flush()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use office_oxide::DocumentIR;
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::*;

    /// The `ir` command must emit the IR's own serde form — not a
    /// projection of it — so every field the IR carries reaches the
    /// output without a consumer to keep in step, and the CLI agrees with
    /// the Python/MCP/C/WASM surfaces byte for byte. Checked on the fields
    /// the old projection missed at one time or another: `speaker_notes`,
    /// `Image::hyperlink`, `Note::marker`/`author`, a `Shape`, span
    /// formatting, conditional formats and data validations.
    #[test]
    fn test_ir_command_emits_the_ir_serde_form_with_every_field() {
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Pptx,
                title: Some("T".into()),
                ..Default::default()
            },
            sections: vec![Section {
                elements: vec![
                    Element::Image(Image {
                        hyperlink: Some("#slide2.xml".to_string()),
                        ..Default::default()
                    }),
                    Element::Shape(Shape {
                        kind: ShapeGeom::Rect,
                        ..Default::default()
                    }),
                    Element::Endnote(Note {
                        id: 0,
                        marker: Some("comment".to_string()),
                        author: Some("Reviewer".to_string()),
                        content: vec![],
                    }),
                    Element::Paragraph(Paragraph {
                        content: vec![InlineContent::Text(TextSpan {
                            underline: Some(UnderlineStyle::Single),
                            font_size_half_pt: Some(36),
                            color: Some([0x12, 0x34, 0x56]),
                            font_name: Some("Calibri".to_string()),
                            vertical_align: Some(VerticalAlign::Superscript),
                            all_caps: true,
                            ..TextSpan::plain("styled")
                        })],
                        alignment: Some(ParagraphAlignment::Center),
                        ..Default::default()
                    }),
                ],
                speaker_notes: Some(vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain("SPEAKER_NOTES_MARKER"))],
                    ..Default::default()
                })]),
                conditional_formats: vec![ConditionalFormat {
                    range: "A1:A10".to_string(),
                    rule_type: "cellIs".to_string(),
                    operator: Some("greaterThan".to_string()),
                    formulas: vec!["100".to_string()],
                }],
                data_validations: vec![DataValidation {
                    range: "B1:B10".to_string(),
                    validation_type: "whole".to_string(),
                    operator: Some("between".to_string()),
                    formula1: Some("1".to_string()),
                    formula2: Some("10".to_string()),
                    allow_blank: true,
                }],
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };

        let mut buf = Vec::new();
        serde_json::to_writer_pretty(&mut buf, &ir).unwrap();
        let rendered = String::from_utf8(buf).unwrap();

        // Same bytes the library-level surfaces produce.
        assert_eq!(rendered, serde_json::to_string_pretty(&ir).unwrap());

        for needle in [
            "SPEAKER_NOTES_MARKER",
            "#slide2.xml",
            r#""type": "shape""#,
            r#""marker": "comment""#,
            r#""author": "Reviewer""#,
            r#""underline": "single""#,
            r#""font_size_half_pt": 36"#,
            r#""font_name": "Calibri""#,
            r#""vertical_align": "superscript""#,
            r#""all_caps": true"#,
            r#""alignment": "center""#,
            r#""rule_type": "cellIs""#,
            r#""validation_type": "whole""#,
        ] {
            assert!(rendered.contains(needle), "missing {needle} in:\n{rendered}");
        }

        // And it round-trips: nothing is projected away.
        let back: DocumentIR = serde_json::from_str(&rendered).unwrap();
        assert_eq!(back, ir);
    }
}
