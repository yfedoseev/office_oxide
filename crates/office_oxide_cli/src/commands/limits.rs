//! `office-oxide limits`: the limits a run reads documents under, and the
//! bounds the library derives from them.

/// Print the limits in force for this run (the defaults, or `--strict` and
/// `--limit` applied to them) and what follows from them, one
/// `name = value` per line; `max` means no bound.
pub fn run(
    options: Option<&office_oxide::options::ParseOptions>,
) -> Result<(), Box<dyn std::error::Error>> {
    use std::io::Write;
    let limits = options.map_or_else(office_oxide::limits::Limits::current, |o| o.limits.clone());
    let value = |v: u64| {
        if v == u64::MAX {
            "max".to_string()
        } else {
            v.to_string()
        }
    };
    let mut out = std::io::stdout().lock();
    writeln!(out, "# limits (set with --limit NAME=VALUE; --strict starts from tighter ones)")?;
    for name in office_oxide::limits::Limits::NAMES {
        let v = limits.get(name).ok_or("a listed limit has no value")?;
        writeln!(out, "{name} = {}", value(v))?;
    }
    writeln!(out, "\n# derived from them")?;
    let derived = serde_json::to_value(limits.derived())?;
    for (name, v) in derived.as_object().ok_or("derived limits are not a map")? {
        let v = v.as_u64().ok_or("a derived limit is not a number")?;
        let shown = if v == u64::MAX || v == usize::MAX as u64 {
            "max".to_string()
        } else {
            v.to_string()
        };
        writeln!(out, "{name} = {shown}")?;
    }
    Ok(())
}
