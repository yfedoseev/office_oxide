//! The size of every type the readers and the IR allocate once per element
//! of a document: a paragraph, run, shape, row, cell, text run or record.
//!
//! A document of a million cells or a hundred thousand runs pays each of
//! these sizes that many times, so a field added inline to one of them is a
//! memory regression proportional to the document. Each bound is the size
//! the type has today; growing one fails here, and the change that grows it
//! must say why the element needs the bytes (or box what is rarely set, as
//! `xlsx::Cell` does). A shrink should lower the bound with it. Run with
//! `--nocapture` for the table.
//!
//! The sizes are those of a 64-bit target; a 32-bit one (wasm32) is smaller.

/// Assert each type is no larger than its bound, and print the table.
macro_rules! sizes {
    ($($ty:ty => $bound:expr),* $(,)?) => {{
        let mut over = Vec::new();
        $(
            let size = std::mem::size_of::<$ty>();
            println!("{size:>5} {:>5}  {}", $bound, stringify!($ty));
            if size > $bound {
                over.push(format!("{} is {size} bytes, bound {}", stringify!($ty), $bound));
            }
        )*
        assert!(over.is_empty(), "{}", over.join("\n"));
    }};
}

#[test]
#[cfg(target_pointer_width = "64")]
fn test_no_per_element_type_grows() {
    sizes! {
        // .docx
        crate::docx::document::BlockElement => 200,
        crate::docx::paragraph::Paragraph => 200,
        crate::docx::paragraph::ParagraphContent => 224,
        crate::docx::paragraph::Run => 224,
        crate::docx::paragraph::RunContent => 80,
        crate::docx::formatting::RunProperties => 200,
        crate::docx::formatting::ParagraphProperties => 176,
        crate::docx::hyperlink::Hyperlink => 96,
        crate::docx::table::TableRow => 40,
        crate::docx::table::TableCell => 120,
        crate::docx::table::TableCellProperties => 96,
        crate::docx::image::DrawingInfo => 248,
        // .pptx
        crate::pptx::shape::Shape => 248,
        crate::pptx::shape::AutoShape => 200,
        crate::pptx::shape::PictureShape => 248,
        crate::pptx::shape::GroupShape => 96,
        crate::pptx::shape::GraphicFrame => 184,
        crate::pptx::shape::TextBody => 24,
        crate::pptx::shape::TextParagraph => 168,
        crate::pptx::shape::TextContent => 176,
        crate::pptx::shape::TextRun => 176,
        crate::pptx::shape::TableRow => 24,
        crate::pptx::shape::TableCell => 40,
        // .xlsx
        crate::xlsx::Cell => 56,
        crate::xlsx::CellValue => 32,
        crate::xlsx::worksheet::Row => 32,
        crate::xlsx::shared_strings::SharedString => 40,
        crate::xlsx::shared_strings::RichTextRun => 112,
        crate::xlsx::styles::CellFormat => 96,
        // .xls
        crate::xls::Cell => 32,
        crate::xls::CellValue => 24,
        crate::xls::XlsComment => 56,
        crate::xls::XlsHyperlink => 32,
        // .doc
        crate::doc::element_types::DocParagraph => 256,
        crate::doc::element_types::FkpParagraph => 32,
        crate::doc::element_types::FkpRun => 32,
        crate::doc::element_types::PapProps => 176,
        crate::doc::element_types::ChpProps => 40,
        crate::doc::element_types::TapCellInfo => 4,
        crate::doc::element_types::Piece => 40,
        // .ppt
        crate::ppt::TextRun => 152,
        crate::ppt::SlideText => 104,
        crate::ppt::CharFormatSpan => 64,
        crate::ppt::ParaFormatSpan => 24,
        crate::ppt::element_types::TableCellData => 32,
        // IR
        crate::ir::Element => 296,
        crate::ir::Paragraph => 296,
        crate::ir::TextSpan => 128,
        crate::ir::InlineContent => 128,
        crate::ir::TableRow => 40,
        crate::ir::TableCell => 304,
        crate::ir::Image => 200,
    }
}
