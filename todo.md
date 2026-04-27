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

### 3. Configure Trusted Publisher on crates.io
**Why:** Eliminates the `CARGO_REGISTRY_TOKEN` long-lived secret. The workflow already uses `rust-lang/crates-io-auth-action` — it just needs the registry side configured.
**Where:** https://crates.io/settings/tokens → "Trusted Publishing" tab

Settings:
- GitHub owner: `yfedoseev`
- Repository: `office_oxide`
- Workflow filename: `release.yml`
- (Optional) Environment name: leave blank

Once configured, delete `CARGO_TOKEN` from https://github.com/yfedoseev/office_oxide/settings/secrets/actions

---

### 4. Configure Trusted Publisher on PyPI
**Why:** Eliminates the `PYPI_API_TOKEN` long-lived secret. The workflow already uses `pypa/gh-action-pypi-publish` with `attestations: true`.
**Where:** https://pypi.org/manage/project/office-oxide/settings/publishing/

Settings:
- GitHub owner: `yfedoseev`
- Repository name: `office_oxide`
- Workflow name: `release.yml`
- Environment name: (leave blank)

Once configured, delete `PYPI_API_TOKEN` from GitHub secrets.

> **First release only:** PyPI Trusted Publishing requires the project to exist first.
> For the very first upload, use a temporary API token, then switch to Trusted Publishing.

---

### 5. Configure Trusted Publisher on npm (×2 packages)
**Why:** `--provenance` in the workflow already generates SLSA attestations; configuring the Trusted Publisher on npm adds the npm-side verification.
**Where:** https://www.npmjs.com/package/office-oxide/access → "Publish access" → "Configure Trusted Publishing"
Repeat for: https://www.npmjs.com/package/office-oxide-wasm/access

Settings for both:
- Repository owner: `yfedoseev`
- Repository name: `office_oxide`
- Workflow filename: `release.yml`

Once configured, `NPM_TOKEN` can optionally be kept as a fallback or removed.

---

### 6. Register for OpenSSF Best Practices Badge
**Why:** Self-assessment gives a `passing` badge, satisfies Scorecard CII-Best-Practices check (5 points), and forces a systematic review of security/quality criteria. Takes ~2 hours.
**Where:** https://www.bestpractices.dev/en/projects/new

After registration:
1. Complete the self-assessment (most criteria are already met — see checklist below)
2. Note your project ID (e.g. `12345`)
3. In `README.md`, uncomment and update the badge line:
   ```
   <!-- [![OpenSSF Best Practices](https://www.bestpractices.dev/projects/NNNN/badge)](https://www.bestpractices.dev/projects/NNNN) -->
   ```
   Replace `NNNN` with the real project ID and remove the comment markers.

**Criteria likely already met:** license, FLOSS, public repo, CONTRIBUTING, CODE_OF_CONDUCT,
SECURITY, CHANGELOG, CI, tests, static analysis, cargo-audit, cargo-deny, docs.rs.

**Criteria to check/add during assessment:**
- Vulnerability response process documented ✅ (SECURITY.md has SLA)
- Test coverage percentage stated ✅ (85% gate in CI)
- No known unpatched CVEs ✅ (cargo-audit in CI)

---

### 7. Rotate NuGet API key annually
**Why:** NuGet does not support OIDC Trusted Publishing yet (as of 2026). The API key expires.
**Where:** https://www.nuget.org/account/apikeys
- Scope the key to `OfficeOxide` only (not all packages)
- Set expiry to 365 days maximum
- Update `NUGET_API_KEY` in https://github.com/yfedoseev/office_oxide/settings/secrets/actions
- Add a calendar reminder for the renewal date

---

## Nice-to-have (post-launch backlog)

- [ ] Add `#![warn(missing_docs)]` to `src/lib.rs` and fix any triggered warnings
- [ ] Add `# Errors` / `# Panics` sections to public Rust API functions
- [ ] Add `cargo-geiger` step to CI (detects unsafe in transitive deps)
- [ ] Add monthly scheduled `cargo outdated` job
- [ ] Populate `docs/architecture/` with module interaction diagrams
- [ ] Add GitHub Discussions category for Q&A (Settings → Features → Discussions)
