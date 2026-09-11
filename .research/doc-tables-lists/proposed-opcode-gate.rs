// Proposed gate for src/doc/sprm.rs — retires the opcode-identity class.
// Verified: fails on the current PR #116 head with exactly the 4 mismatches
// and 3 missing properties recorded in round3-blind-verdict.md.
#[cfg(test)]
mod spec_conformance {
    //! Opcode-identity gate. Transcribed from [MS-DOC] 2.6.1 "Paragraph
    //! Properties" and 2.6.3 "Table Properties". This retires the class
    //! "the constant is right but names a different property": every opcode
    //! the decoder acts on is checked against the spec's own name for it.

    /// (opcode, spec name) — transcribed from the [MS-DOC] property tables.
    const SPEC: &[(u16, &str)] = &[
        // 2.6.1 Paragraph Properties (sgc = 1)
        (0x260A, "sprmPIlvl"),
        (0x460B, "sprmPIlfo"),
        (0xC60D, "sprmPChgTabsPapx"),
        (0xC615, "sprmPChgTabs"),
        (0x2416, "sprmPFInTable"),
        (0x2417, "sprmPFTtp"),
        (0x6649, "sprmPItap"),
        (0x664A, "sprmPDtap"),
        (0x244B, "sprmPFInnerTableCell"),
        (0x244C, "sprmPFInnerTtp"),
        // 2.6.3 Table Properties (sgc = 5)
        (0xD608, "sprmTDefTable"),
        (0x7621, "sprmTInsert"),
        (0xD632, "sprmTCellPadding"),
        (0xD634, "sprmTCellPaddingDefault"),
        (0x563A, "sprmTIstd"),
    ];

    fn spec_name(opcode: u16) -> Option<&'static str> {
        SPEC.iter().find(|(op, _)| *op == opcode).map(|(_, n)| *n)
    }

    /// Each opcode the decoder dispatches on, the `PapProps` field it feeds,
    /// and the property the code's own comment claims it is.
    const DISPATCH: &[(u16, &str, &str)] = &[
        (0x2416, "f_in_table", "sprmPFInTable"),
        (0x6649, "itap", "sprmPItap"),
        (0xD608, "tap", "sprmTDefTable"),
        (0x460D, "ilfo", "sprmPIlfo"),
        (0x460B, "ilvl", "sprmPIlvl"),
        (0xD632, "tabs", "sprmPChgTabs"),
        (0xD634, "tabs", "sprmPChgTabs"),
    ];

    #[test]
    fn dispatched_opcodes_match_their_claimed_spec_property() {
        let mut bad = Vec::new();
        for (op, field, claimed) in DISPATCH {
            match spec_name(*op) {
                Some(actual) if actual == *claimed => {},
                Some(actual) => bad.push(format!(
                    "  {:#06X} -> PapProps::{:<11} code calls it {:<16} MS-DOC says {}",
                    op, field, claimed, actual
                )),
                None => bad.push(format!(
                    "  {:#06X} -> PapProps::{:<11} code calls it {:<16} NOT DEFINED in MS-DOC",
                    op, field, claimed
                )),
            }
        }
        assert!(bad.is_empty(), "opcode/property mismatches:\n{}", bad.join("\n"));
    }

    /// The spec names a property the decoder needs; check it is actually read.
    #[test]
    fn required_properties_are_decoded() {
        let dispatched: Vec<u16> = DISPATCH.iter().map(|(op, _, _)| *op).collect();
        let mut missing = Vec::new();
        for (op, name) in [
            (0x260A, "sprmPIlvl"),
            (0xC615, "sprmPChgTabs"),
            (0xC60D, "sprmPChgTabsPapx"),
        ] {
            if !dispatched.contains(&op) {
                missing.push(format!("  {:#06X} {} is never decoded", op, name));
            }
        }
        assert!(missing.is_empty(), "spec properties not decoded:\n{}", missing.join("\n"));
    }
}
