#![no_main]

// The WPS downrev packages sit behind a valid CFB container, a valid record
// tree and a valid ZIP holding `drs/shapexml.xml`; byte-level fuzzing of whole
// files almost never gets that far, and when it does the package itself has to
// be a well-formed archive. This target builds a well-formed deck in code and
// hands the fuzzer's bytes in as the carrier payloads, so every run reaches the
// tertiary property-table walk, the declared-length bound, the fallback
// signature scan and the package reader. None may panic, overflow or hang; a
// malformed package must be skipped, and the per-slide package cap must bound
// the work.
#[path = "../../tests/common/mod.rs"]
mod common;

use std::io::Cursor;

use common::{
    downrev_group_from_carrier, downrev_shape_group, downrev_shape_xml, downrev_zip, ppt_record,
    ppt_stream_with_persist, RT_DOWNREV_CARRIER,
};
use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

fuzz_target!(|data: &[u8]| {
    // byte 0: carrier count (1-4); the rest is split into that many carrier
    // payloads, one shape group each, on a single slide.
    let Some((&count, rest)) = data.split_first() else {
        return;
    };
    let n = usize::from(count % 4) + 1;
    let chunk = rest.len() / n;
    let mut groups: Vec<Vec<u8>> = (0..n)
        .map(|i| {
            let payload = &rest[i * chunk..((i + 1) * chunk).min(rest.len())];
            downrev_group_from_carrier(ppt_record(RT_DOWNREV_CARRIER, 1, 3, payload))
        })
        .collect();
    // One well-formed package alongside the fuzzed ones, so the extraction
    // path stays exercised in every run.
    let good = downrev_shape_group(&downrev_zip(&downrev_shape_xml(&["Fuzz heading", "Fuzz body"])));
    groups.push(good);
    let master = vec![downrev_shape_group(&downrev_zip(&downrev_shape_xml(&[
        "Master title style",
    ])))];
    let bytes = ppt_stream_with_persist(&[groups], &master);
    if let Ok(doc) = Document::from_reader(Cursor::new(bytes), DocumentFormat::Ppt) {
        let _ = doc.plain_text();
        let _ = doc.to_markdown();
        let _ = doc.to_ir();
    }
});
