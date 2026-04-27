# office_oxide — Manual Release Checklist

Everything automated has been committed. The items below require browser/account access
and cannot be done from the terminal.

---

## Before first public release

### 1. Enable GitHub Private Vulnerability Reporting
**Why:** Lets reporters file a private draft advisory without sending email. Scorecard bonus points.
**Where:** https://github.com/yfedoseev/office_oxide/settings/security_analysis
→ "Private vulnerability reporting" → **Enable**

---

### 2. Set branch protection rules on `main`
**Why:** Scorecard Branch-Protection check (High). Prevents force-pushes and bypassed CI.
**Where:** https://github.com/yfedoseev/office_oxide/settings/branches → Add rule for `main`

Recommended settings:
- [x] Require a pull request before merging
  - [x] Require approvals: 1
  - [x] Dismiss stale reviews when new commits are pushed
- [x] Require status checks to pass before merging
  - Required checks: `fmt`, `clippy`, `test (ubuntu-latest, stable)`, `audit`, `deny`, `dco`
- [x] Require branches to be up to date before merging
- [x] Do not allow bypassing the above settings
- [ ] Restrict force-pushes (leave off for solo maintainer convenience, enable when team grows)

---

## Nice-to-have (post-launch backlog)

- [ ] Add `#![warn(missing_docs)]` to `src/lib.rs` and fix any triggered warnings
- [ ] Add `# Errors` / `# Panics` sections to public Rust API functions
- [ ] Add `cargo-geiger` step to CI (detects unsafe in transitive deps)
- [ ] Add monthly scheduled `cargo outdated` job
- [ ] Populate `docs/architecture/` with module interaction diagrams
- [ ] Add GitHub Discussions category for Q&A (Settings → Features → Discussions)
