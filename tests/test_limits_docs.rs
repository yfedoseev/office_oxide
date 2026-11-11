//! The README's table of limits states the defaults the library uses.

use office_oxide::limits::Limits;

/// `1 GiB`, `512 MiB`, `256 Mi`, `100,000`, `3`.
fn shown(value: u64, bytes: bool) -> String {
    const GI: u64 = 1 << 30;
    const MI: u64 = 1 << 20;
    if value >= MI && value.is_multiple_of(GI) {
        format!("{} Gi{}", value / GI, if bytes { "B" } else { "" })
    } else if value >= MI && value.is_multiple_of(MI) {
        format!("{} Mi{}", value / MI, if bytes { "B" } else { "" })
    } else {
        let digits = value.to_string();
        let mut out = String::new();
        for (i, c) in digits.chars().enumerate() {
            if i > 0 && (digits.len() - i).is_multiple_of(3) {
                out.push(',');
            }
            out.push(c);
        }
        out
    }
}

#[test]
fn test_readme_limits_table_states_the_defaults() {
    let readme =
        std::fs::read_to_string(concat!(env!("CARGO_MANIFEST_DIR"), "/README.md")).unwrap();
    let defaults = Limits::default();
    for name in Limits::NAMES {
        let row = readme
            .lines()
            .find(|l| l.starts_with(&format!("| `{name}` |")))
            .unwrap_or_else(|| panic!("README's limits table has no row for `{name}`"));
        let cell = row.split('|').nth(2).unwrap().trim();
        let value = defaults.get(name).unwrap();
        let expected = shown(value, name.ends_with("_bytes"));
        assert_eq!(cell, expected, "README states `{name}` = {cell}, the default is {value}");
    }
}
