//! BIFF8 record types and low-level record iterator.

#![allow(dead_code)]

use super::error::Result;

// ── Record type IDs ──
pub const RT_BOF: u16 = 0x0809;
pub const RT_EOF: u16 = 0x000A;
pub const RT_BOUNDSHEET: u16 = 0x0085;
pub const RT_SST: u16 = 0x00FC;
pub const RT_CONTINUE: u16 = 0x003C;
pub const RT_LABELSST: u16 = 0x00FD;
pub const RT_NUMBER: u16 = 0x0203;
pub const RT_RK: u16 = 0x027E;
pub const RT_MULRK: u16 = 0x00BD;
pub const RT_FORMULA: u16 = 0x0006;
pub const RT_BOOLERR: u16 = 0x0205;
pub const RT_LABEL: u16 = 0x0204;
pub const RT_RSTRING: u16 = 0x00D6;
pub const RT_BLANK: u16 = 0x0201;
pub const RT_MULBLANK: u16 = 0x00BE;
pub const RT_FORMAT: u16 = 0x041E;
pub const RT_XF: u16 = 0x00E0;
pub const RT_DATEMODE: u16 = 0x0022;
pub const RT_CODEPAGE: u16 = 0x0042;
pub const RT_FILEPASS: u16 = 0x002F;
pub const RT_STRING: u16 = 0x0207;
/// Merged cell ranges for the current worksheet ([MS-XLS] §2.4.180).
pub const RT_MERGEDCELLS: u16 = 0x00E5;
pub const RT_NAME: u16 = 0x0018;
pub const RT_EXTERNSHEET: u16 = 0x0017;
pub const RT_SUPBOOK: u16 = 0x01AE;
/// Chart series/trendline/axis/chart title text ([MS-XLS] §2.4.254), found
/// inside a chart's nested `BOF..EOF` substream (issue #246).
pub const RT_SERIESTEXT: u16 = 0x100D;
/// Begins a conditional-formatting rule group: the cell range(s) it
/// applies to, plus how many `CF` records follow ([MS-XLS] §2.4.56,
/// record type 432 = 0x1B0). Issue #252.
pub const RT_CONDFMT: u16 = 0x01B0;
/// One conditional-formatting rule within the group opened by the
/// preceding `CONDFMT` ([MS-XLS] §2.4.42, record type 433 = 0x1B1).
pub const RT_CF: u16 = 0x01B1;

/// One data validation rule ([MS-XLS] §2.4.44, record type 446 = 0x1BE).
/// Issue #275.
pub const RT_DV: u16 = 0x01BE;

/// A cell (or cell range's) hyperlink target ([MS-XLS] §2.4.130, record
/// type 440 = 0x1B8). Issue #306.
pub const RT_HLINK: u16 = 0x01B8;

/// Declares a drawing shape/object and its object id ([MS-XLS] §2.4.181,
/// record type 93 = 0x5D). Only used here to correlate a `NOTE`'s
/// `shapeid` to its comment text (issue #307).
pub const RT_OBJ: u16 = 0x005D;
/// A cell comment's position and author ([MS-XLS] §2.4.178, record type
/// 28 = 0x1C). Issue #307.
pub const RT_NOTE: u16 = 0x001C;
/// The text of the drawing shape declared by the immediately preceding
/// `OBJ` record ([MS-XLS] §2.4.326, record type 438 = 0x1B6). Issue #307.
pub const RT_TXO: u16 = 0x01B6;

/// A raw BIFF record: type + data (may span CONTINUE records).
#[derive(Debug, Clone)]
pub struct BiffRecord {
    pub record_type: u16,
    pub data: Vec<u8>,
}

/// Iterate over BIFF records in a byte stream, merging CONTINUE records.
pub struct RecordIter<'a> {
    data: &'a [u8],
    pos: usize,
    last_type: u16,
}

impl<'a> RecordIter<'a> {
    pub fn new(data: &'a [u8]) -> Self {
        Self {
            data,
            pos: 0,
            last_type: 0,
        }
    }

    /// Read the next raw record (without merging CONTINUE).
    fn read_raw(&mut self) -> Option<Result<(u16, Vec<u8>)>> {
        if self.pos + 4 > self.data.len() {
            return None;
        }

        let rt = u16::from_le_bytes([self.data[self.pos], self.data[self.pos + 1]]);
        let size = u16::from_le_bytes([self.data[self.pos + 2], self.data[self.pos + 3]]) as usize;
        self.pos += 4;

        if self.pos + size > self.data.len() {
            // Truncated record — use what's available instead of erroring.
            let available = self.data.len() - self.pos;
            let data = self.data[self.pos..self.pos + available].to_vec();
            self.pos = self.data.len();
            return Some(Ok((rt, data)));
        }

        let data = self.data[self.pos..self.pos + size].to_vec();
        self.pos += size;
        Some(Ok((rt, data)))
    }
}

impl<'a> Iterator for RecordIter<'a> {
    type Item = Result<BiffRecord>;

    fn next(&mut self) -> Option<Self::Item> {
        let (rt, mut data) = match self.read_raw()? {
            Ok(v) => v,
            Err(e) => return Some(Err(e)),
        };

        self.last_type = rt;

        // Merge subsequent CONTINUE records into this record's data.
        loop {
            if self.pos + 4 > self.data.len() {
                break;
            }
            let next_rt = u16::from_le_bytes([self.data[self.pos], self.data[self.pos + 1]]);
            if next_rt != RT_CONTINUE {
                break;
            }
            // Consume the CONTINUE record.
            match self.read_raw() {
                Some(Ok((_rt, cont_data))) => data.extend_from_slice(&cont_data),
                Some(Err(e)) => return Some(Err(e)),
                None => break,
            }
        }

        Some(Ok(BiffRecord {
            record_type: rt,
            data,
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_record(rt: u16, data: &[u8]) -> Vec<u8> {
        let mut buf = Vec::new();
        buf.extend_from_slice(&rt.to_le_bytes());
        buf.extend_from_slice(&(data.len() as u16).to_le_bytes());
        buf.extend_from_slice(data);
        buf
    }

    #[test]
    fn iterate_single_record() {
        let stream = make_record(RT_BOF, &[0x00, 0x06, 0x05, 0x00]);
        let records: Vec<_> = RecordIter::new(&stream)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(records.len(), 1);
        assert_eq!(records[0].record_type, RT_BOF);
        assert_eq!(records[0].data, &[0x00, 0x06, 0x05, 0x00]);
    }

    #[test]
    fn iterate_multiple_records() {
        let mut stream = make_record(RT_BOF, &[0x00, 0x06]);
        stream.extend(make_record(RT_EOF, &[]));
        let records: Vec<_> = RecordIter::new(&stream)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(records.len(), 2);
        assert_eq!(records[0].record_type, RT_BOF);
        assert_eq!(records[1].record_type, RT_EOF);
    }

    #[test]
    fn continue_records_merged() {
        let mut stream = make_record(RT_SST, &[0x01, 0x02]);
        stream.extend(make_record(RT_CONTINUE, &[0x03, 0x04]));
        stream.extend(make_record(RT_CONTINUE, &[0x05]));
        stream.extend(make_record(RT_EOF, &[]));
        let records: Vec<_> = RecordIter::new(&stream)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(records.len(), 2); // SST (merged) + EOF
        assert_eq!(records[0].record_type, RT_SST);
        assert_eq!(records[0].data, &[0x01, 0x02, 0x03, 0x04, 0x05]);
    }

    #[test]
    fn truncated_record_tolerant() {
        // Record says 10 bytes but only 2 available — should return partial data.
        let mut stream = Vec::new();
        stream.extend_from_slice(&RT_BOF.to_le_bytes());
        stream.extend_from_slice(&10u16.to_le_bytes());
        stream.extend_from_slice(&[0x00, 0x00]);
        let results: Vec<_> = RecordIter::new(&stream)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(results.len(), 1);
        assert_eq!(results[0].record_type, RT_BOF);
        assert_eq!(results[0].data.len(), 2); // only 2 bytes available
    }

    #[test]
    fn empty_stream() {
        let records: Vec<_> = RecordIter::new(&[])
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert!(records.is_empty());
    }
}
