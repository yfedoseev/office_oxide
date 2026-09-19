//! Text extraction from DrawingML chart parts (`c:chartSpace`).
//!
//! A chart embedded in a DOCX (`word/charts/chartN.xml`) or a PPTX
//! (`ppt/charts/chartN.xml`) keeps all of its human-meaningful words in a
//! *separate* part, reachable only through the drawing's
//! `<c:chart r:id="…"/>` relationship. The referencing part itself holds no
//! text at all, so a reader that never opens the chart part loses the chart
//! title, the axis titles, every category label and every cached data value.
//!
//! This module holds the reader for that part. It is deliberately
//! independent of the XLSX reader's own `extract_chart_text` (which does the
//! same job for `xl/charts/chartN.xml`); both walk the identical `c:` schema,
//! but keeping them separate avoids coupling two formats' output shapes
//! together.
//!
//! Output shape (one line per element, matching what the XLSX side emits so
//! that chart text reads the same whichever format carried the chart):
//!
//! ```text
//! Title: Dollars per Group
//! Categories: Group 1, Group 2
//! Series 1: 15.53, 27.32
//! ```

/// Upper bounds on what we will accumulate out of one chart part. A chart is
/// decoration: no legitimate one needs more than this, and without caps a
/// hostile part (millions of `<c:pt>` entries, a megabyte-long title) turns
/// text extraction into a memory-exhaustion vector.
const MAX_SERIES: usize = 512;
const MAX_VALUES_PER_SERIES: usize = 4096;
const MAX_CATEGORIES: usize = 4096;
const MAX_TITLE_LEN: usize = 4096;
const MAX_CELL_LEN: usize = 512;

#[derive(Default)]
struct Series {
    name: String,
    values: Vec<String>,
    /// `<c:bubbleSize>` cached values. Kept apart from `values` so a bubble
    /// chart's radii don't read as extra plotted values.
    sizes: Vec<String>,
}

/// Extract the readable text of a DrawingML chart part as a list of lines.
///
/// Returns an empty vector when the part carries no title, no categories and
/// no series — i.e. when there is nothing a reader would have seen.
pub fn chart_text_lines(xml: &[u8]) -> Vec<String> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().check_end_names = false;
    reader.config_mut().check_comments = false;
    let mut buf = Vec::new();

    // Local names of the open elements, innermost last. Pushed on Start and
    // popped on End, so on an End event the stack already describes the
    // *parent* scope of the element that just closed.
    let mut stack: Vec<String> = Vec::new();
    // How deep we are inside `<c:title>` elements (chart title and axis
    // titles both use it). Rich text only counts as a title when > 0.
    let mut title_depth: u32 = 0;
    let mut title_buf = String::new();
    let mut titles: Vec<String> = Vec::new();

    let mut series: Vec<Series> = Vec::new();
    let mut cur_series: Option<Series> = None;
    let mut cur_cats: Vec<String> = Vec::new();
    let mut shared_categories: Vec<String> = Vec::new();
    // Text of the `<c:v>` currently open. Accumulated untrimmed so a value
    // split across several Text / entity-reference events keeps its interior
    // spacing; it is trimmed once, on close.
    let mut cur_v = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.local_name().as_ref().to_string();
                match local.as_str() {
                    "ser" => {
                        cur_series = Some(Series::default());
                        cur_cats.clear();
                    },
                    "title" => title_depth += 1,
                    _ => {},
                }
                stack.push(local);
            },
            Ok(quick_xml::events::Event::End(e)) => {
                let local = e.local_name().as_ref().to_string();
                stack.pop();
                match local.as_str() {
                    // A title's runs are grouped into `<a:p>` paragraphs;
                    // without a separator "Sales" + "2024" would fuse.
                    "p" if title_depth > 0 => {
                        if !title_buf.is_empty() && !title_buf.ends_with(' ') {
                            title_buf.push(' ');
                        }
                    },
                    "title" => {
                        title_depth = title_depth.saturating_sub(1);
                        let t = title_buf.trim();
                        if !t.is_empty() {
                            titles.push(t.to_string());
                        }
                        title_buf.clear();
                    },
                    "v" => {
                        let val = cur_v.trim().to_string();
                        cur_v.clear();
                        if val.is_empty() {
                            continue;
                        }
                        // `<c:tx><c:v>` / `<c:tx><c:strRef><c:strCache>…`
                        // inside a `<c:title>` is a title pulled from a cell
                        // rather than typed in place.
                        if title_depth > 0 && cur_series.is_none() {
                            push_capped(&mut title_buf, &val);
                            continue;
                        }
                        let Some(s) = cur_series.as_mut() else {
                            continue;
                        };
                        let scope = |tag: &str| stack.iter().any(|t| t == tag);
                        if scope("tx") {
                            if s.name.is_empty() {
                                s.name = val;
                            }
                        } else if scope("cat") || scope("xVal") {
                            // Scatter charts have no `<c:cat>`; their x
                            // values play the same role.
                            if cur_cats.len() < MAX_CATEGORIES {
                                cur_cats.push(val);
                            }
                        } else if scope("bubbleSize") {
                            if s.sizes.len() < MAX_VALUES_PER_SERIES {
                                s.sizes.push(val);
                            }
                        } else if scope("val") || scope("yVal") {
                            if s.values.len() < MAX_VALUES_PER_SERIES {
                                s.values.push(val);
                            }
                        }
                    },
                    "ser" => {
                        if let Some(mut s) = cur_series.take() {
                            // Every series in a chart normally repeats the
                            // same category list; keep the first non-empty
                            // one and print it once.
                            if shared_categories.is_empty() && !cur_cats.is_empty() {
                                shared_categories = std::mem::take(&mut cur_cats);
                            } else {
                                cur_cats.clear();
                            }
                            if s.name.is_empty() {
                                s.name = format!("Series {}", series.len() + 1);
                            }
                            if series.len() < MAX_SERIES {
                                series.push(s);
                            }
                        }
                    },
                    _ => {},
                }
            },
            Ok(quick_xml::events::Event::Text(t)) => {
                if let Ok(s) = crate::core::xml::unescape_text(&t) {
                    append_scoped(&mut stack, title_depth, &s, &mut title_buf, &mut cur_v);
                }
            },
            // Entity references arrive as their own event: without folding
            // them in here, `AT&amp;T` in a chart title would read "ATT".
            Ok(quick_xml::events::Event::GeneralRef(ref r)) => {
                if let Ok(s) = crate::core::xml::resolve_general_ref(r) {
                    append_scoped(&mut stack, title_depth, &s, &mut title_buf, &mut cur_v);
                }
            },
            Ok(quick_xml::events::Event::CData(t)) => {
                let s = t.to_string();
                append_scoped(&mut stack, title_depth, &s, &mut title_buf, &mut cur_v);
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
        buf.clear();
    }

    let mut lines = Vec::new();
    if !titles.is_empty() {
        lines.push(format!("Title: {}", titles.join(" — ")));
    }
    if !shared_categories.is_empty() {
        lines.push(format!("Categories: {}", shared_categories.join(", ")));
    }
    for s in &series {
        if s.values.is_empty() {
            lines.push(format!("Series: {}", s.name));
        } else {
            lines.push(format!("{}: {}", s.name, s.values.join(", ")));
        }
        if !s.sizes.is_empty() {
            lines.push(format!("{} (bubble size): {}", s.name, s.sizes.join(", ")));
        }
    }
    lines
}

/// Extract the readable text of a chart part as a single newline-separated
/// string. Empty when the chart carries no text at all.
pub fn extract_chart_text(xml: &[u8]) -> String {
    chart_text_lines(xml).join("\n")
}

/// Route a text fragment to the title buffer or the current `<c:v>` buffer,
/// depending on which element is currently open.
fn append_scoped(
    stack: &mut [String],
    title_depth: u32,
    s: &str,
    title_buf: &mut String,
    cur_v: &mut String,
) {
    match stack.last().map(|v| v.as_str()) {
        // `<a:t>` is DrawingML rich text. It appears in data labels too, so
        // only count it when a `<c:title>` is open.
        Some("t") if title_depth > 0 => push_capped(title_buf, s),
        Some("v") => {
            if cur_v.len() < MAX_CELL_LEN {
                cur_v.push_str(s);
            }
        },
        _ => {},
    }
}

fn push_capped(dst: &mut String, s: &str) {
    if dst.len() < MAX_TITLE_LEN {
        dst.push_str(s);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_chart_text_title_categories_series() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Dollars per Group</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>Group 1</c:v></c:pt>
            <c:pt idx="1"><c:v>Group 2</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>15.53</c:v></c:pt>
            <c:pt idx="1"><c:v>27.32</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;
        let lines = chart_text_lines(xml);
        assert_eq!(
            lines,
            vec![
                "Title: Dollars per Group".to_string(),
                "Categories: Group 1, Group 2".to_string(),
                "Revenue: 15.53, 27.32".to_string(),
            ]
        );
    }

    #[test]
    fn test_chart_text_scatter_uses_xval_yval() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:scatterChart><c:ser>
    <c:xVal><c:numRef><c:numCache>
      <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt>
    </c:numCache></c:numRef></c:xVal>
    <c:yVal><c:numRef><c:numCache>
      <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt>
    </c:numCache></c:numRef></c:yVal>
  </c:ser></c:scatterChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let lines = chart_text_lines(xml);
        assert_eq!(lines, vec!["Categories: 1, 2", "Series 1: 10, 20"]);
    }

    #[test]
    fn test_chart_text_multi_run_title_keeps_spacing() {
        // A title typed as two runs must not fuse into "SalesReport", and
        // two paragraphs must not fuse either.
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:title><c:tx><c:rich>
    <a:p><a:r><a:t>Sales </a:t></a:r><a:r><a:t>Report</a:t></a:r></a:p>
    <a:p><a:r><a:t>2024</a:t></a:r></a:p>
  </c:rich></c:tx></c:title></c:chart>
</c:chartSpace>"#;
        assert_eq!(extract_chart_text(xml), "Title: Sales Report 2024");
    }

    #[test]
    fn test_chart_text_entity_in_title_survives() {
        let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>AT&amp;T</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart>
</c:chartSpace>"#;
        assert_eq!(extract_chart_text(xml), "Title: AT&T");
    }

    #[test]
    fn test_chart_text_empty_when_no_content() {
        let xml =
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea/></c:chart></c:chartSpace>"#;
        assert!(chart_text_lines(xml).is_empty());
        assert!(extract_chart_text(xml).is_empty());
    }

    #[test]
    fn test_chart_text_axis_titles_included() {
        let xml = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Growth</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:catAx><c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>
      <c:valAx><c:title><c:tx><c:rich><a:p><a:r><a:t>Units</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;
        assert_eq!(extract_chart_text(xml), "Title: Growth — Quarter — Units");
    }
}
