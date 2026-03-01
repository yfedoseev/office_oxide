use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::NsReader;

use office_core::xml;

use crate::formatting::Justification;

/// All numbering definitions from `word/numbering.xml`.
#[derive(Debug, Clone, Default)]
pub struct NumberingDefinitions {
    pub abstract_nums: HashMap<u32, AbstractNum>,
    pub instances: HashMap<u32, NumberingInstance>,
}

/// An abstract numbering definition.
#[derive(Debug, Clone)]
pub struct AbstractNum {
    pub abstract_num_id: u32,
    pub levels: HashMap<u8, NumberingLevel>,
}

/// A single level within a numbering definition.
#[derive(Debug, Clone)]
pub struct NumberingLevel {
    pub start: u32,
    pub format: NumberFormat,
    pub level_text: String,
    pub justification: Option<Justification>,
}

/// A concrete numbering instance referencing an abstract definition.
#[derive(Debug, Clone)]
pub struct NumberingInstance {
    pub num_id: u32,
    pub abstract_num_id: u32,
    pub overrides: HashMap<u8, NumberingLevel>,
}

/// Number format type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NumberFormat {
    Decimal,
    Bullet,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
    None,
    Other(String),
}

impl NumberingDefinitions {
    /// Parse `word/numbering.xml` content.
    pub fn parse(xml_data: &[u8]) -> office_core::Result<Self> {
        let mut reader = xml::make_reader(xml_data);
        let wml = xml::ns::WML;
        let mut defs = NumberingDefinitions::default();

        loop {
            match reader.read_resolved_event()? {
                (ref resolve, Event::Start(ref e)) => {
                    if xml::matches_ns(resolve, wml) {
                        match e.local_name().as_ref() {
                            b"abstractNum" => {
                                if let Some(an) = parse_abstract_num(&mut reader, e)? {
                                    defs.abstract_nums.insert(an.abstract_num_id, an);
                                }
                            }
                            b"num" => {
                                if let Some(inst) = parse_num_instance(&mut reader, e)? {
                                    defs.instances.insert(inst.num_id, inst);
                                }
                            }
                            _ => {}
                        }
                    }
                }
                (_, Event::Eof) => break,
                _ => {}
            }
        }
        Ok(defs)
    }

    /// Resolve a numbering level for a given numId + ilvl.
    pub fn resolve_level(&self, num_id: u32, ilvl: u8) -> Option<&NumberingLevel> {
        let instance = self.instances.get(&num_id)?;
        if let Some(level) = instance.overrides.get(&ilvl) {
            return Some(level);
        }
        let abstract_num = self.abstract_nums.get(&instance.abstract_num_id)?;
        abstract_num.levels.get(&ilvl)
    }
}

fn parse_abstract_num(
    reader: &mut NsReader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> office_core::Result<Option<AbstractNum>> {
    let wml = xml::ns::WML;
    let abstract_num_id = match xml::optional_attr_str(start, b"w:abstractNumId")? {
        Some(id) => id.parse().unwrap_or(0),
        None => return Ok(None),
    };
    let mut levels = HashMap::new();

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) => {
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"lvl" {
                    let ilvl = xml::optional_attr_str(e, b"w:ilvl")?
                        .and_then(|v| v.parse::<u8>().ok())
                        .unwrap_or(0);
                    let level = parse_numbering_level(reader)?;
                    levels.insert(ilvl, level);
                } else {
                    xml::skip_element(reader)?;
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"abstractNum" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    Ok(Some(AbstractNum {
        abstract_num_id,
        levels,
    }))
}

fn parse_numbering_level(reader: &mut NsReader<&[u8]>) -> office_core::Result<NumberingLevel> {
    let wml = xml::ns::WML;
    let mut start_val = 1u32;
    let mut format = NumberFormat::Decimal;
    let mut level_text = String::new();
    let mut justification: Option<Justification> = None;

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) | (ref resolve, Event::Empty(ref e)) => {
                if xml::matches_ns(resolve, wml) {
                    match e.local_name().as_ref() {
                        b"start" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                start_val = val.parse().unwrap_or(1);
                            }
                        }
                        b"numFmt" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                format = parse_number_format(&val);
                            }
                        }
                        b"lvlText" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                level_text = val.into_owned();
                            }
                        }
                        b"lvlJc" => {
                            if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                                justification =
                                    Some(crate::formatting::parse_justification_value(&val));
                            }
                        }
                        b"pPr" | b"rPr" => {
                            // Skip sub-properties for now (they apply to the numbering marker)
                        }
                        _ => {}
                    }
                }
                // If this was a Start event, we need to skip its content
                // But since we match both Start and Empty, we can't easily distinguish.
                // Let's handle it differently.
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"lvl" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    Ok(NumberingLevel {
        start: start_val,
        format,
        level_text,
        justification,
    })
}

fn parse_number_format(val: &str) -> NumberFormat {
    match val {
        "decimal" => NumberFormat::Decimal,
        "bullet" => NumberFormat::Bullet,
        "lowerLetter" => NumberFormat::LowerLetter,
        "upperLetter" => NumberFormat::UpperLetter,
        "lowerRoman" => NumberFormat::LowerRoman,
        "upperRoman" => NumberFormat::UpperRoman,
        "none" => NumberFormat::None,
        other => NumberFormat::Other(other.to_string()),
    }
}

fn parse_num_instance(
    reader: &mut NsReader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> office_core::Result<Option<NumberingInstance>> {
    let wml = xml::ns::WML;
    let num_id = match xml::optional_attr_str(start, b"w:numId")? {
        Some(id) => id.parse().unwrap_or(0),
        None => return Ok(None),
    };
    let mut abstract_num_id = 0u32;
    let overrides = HashMap::new();

    loop {
        match reader.read_resolved_event()? {
            (ref resolve, Event::Start(ref e)) | (ref resolve, Event::Empty(ref e)) => {
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"abstractNumId" {
                    if let Ok(Some(val)) = xml::optional_attr_str(e, b"w:val") {
                        abstract_num_id = val.parse().unwrap_or(0);
                    }
                }
            }
            (ref resolve, Event::End(ref e)) => {
                if xml::matches_ns(resolve, wml) && e.local_name().as_ref() == b"num" {
                    break;
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }

    Ok(Some(NumberingInstance {
        num_id,
        abstract_num_id,
        overrides,
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_NUMBERING: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#61623;"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%2."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"#;

    #[test]
    fn parse_numbering_defs() {
        let defs = NumberingDefinitions::parse(SAMPLE_NUMBERING).unwrap();
        assert_eq!(defs.abstract_nums.len(), 1);
        assert_eq!(defs.instances.len(), 1);

        let an = defs.abstract_nums.get(&0).unwrap();
        assert_eq!(an.levels.len(), 2);

        let inst = defs.instances.get(&1).unwrap();
        assert_eq!(inst.abstract_num_id, 0);
    }

    #[test]
    fn resolve_numbering_level() {
        let defs = NumberingDefinitions::parse(SAMPLE_NUMBERING).unwrap();
        let level = defs.resolve_level(1, 0).unwrap();
        assert_eq!(level.format, NumberFormat::Bullet);
        assert_eq!(level.start, 1);

        let level1 = defs.resolve_level(1, 1).unwrap();
        assert_eq!(level1.format, NumberFormat::Decimal);
    }
}
