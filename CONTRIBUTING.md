# Contributing to office_oxide

Thank you for your interest in contributing! This document provides guidelines and information for contributors.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [Project Structure](#project-structure)
- [Development Workflow](#development-workflow)
- [Coding Standards](#coding-standards)
- [Testing](#testing)
- [Documentation](#documentation)
- [Submitting Changes](#submitting-changes)
- [License](#license)

## Code of Conduct

This project adheres to the [Contributor Covenant Code of Conduct](CODE_OF_CONDUCT.md). By participating, you are expected to uphold this code. Please report unacceptable behavior by opening an issue or contacting the maintainers.

## Getting Started

### Prerequisites

- **Rust**: 1.85+ ([Install Rust](https://rustup.rs/))
- **Python**: 3.8+ (for Python bindings)
- **Git**: For version control

### Optional Tools

- **cargo-watch**: Auto-reload on file changes
  ```bash
  cargo install cargo-watch
  ```

- **cargo-llvm-cov**: Code coverage
  ```bash
  cargo install cargo-llvm-cov
  ```

- **maturin**: Python packaging
  ```bash
  pip install maturin
  # or
  uv tool install maturin
  ```

- **pre-commit**: Git hooks
  ```bash
  pip install pre-commit
  pre-commit install
  ```

## Development Setup

1. **Fork and clone** the repository:
   ```bash
   git clone https://github.com/YOUR_USERNAME/office_oxide.git
   cd office_oxide
   ```

2. **Build the project**:
   ```bash
   cargo build
   ```

3. **Run tests**:
   ```bash
   cargo test
   ```

4. **Set up pre-commit hooks** (recommended):
   ```bash
   pre-commit install
   ```

## Project Structure

```
office_oxide/
├── src/
│   ├── lib.rs             # Unified Document API + convenience functions
│   ├── core/              # Shared OPC/ZIP/XML/theme primitives (55 tests)
│   ├── cfb/               # CFBF/OLE2 container reader (18 tests)
│   ├── docx/              # Word document (.docx) — read/write/edit (36 tests)
│   ├── xlsx/              # Excel spreadsheet (.xlsx) — read/write/edit (57 tests)
│   ├── pptx/              # PowerPoint presentation (.pptx) — read/write/edit (40 tests)
│   ├── doc/               # Legacy Word Binary (.doc) (15 tests)
│   ├── xls/               # Legacy Excel Binary (.xls) (24 tests)
│   ├── ppt/               # Legacy PowerPoint Binary (.ppt) (15 tests)
│   ├── ir.rs              # Format-agnostic DocumentIR
│   ├── ir_render.rs       # IR → plain_text / markdown / html
│   ├── create.rs          # IR → DOCX/XLSX/PPTX creation
│   ├── edit.rs            # Unified EditableDocument API
│   ├── python.rs          # PyO3 bindings (feature = python)
│   ├── wasm.rs            # wasm-bindgen bindings (feature = wasm)
│   └── ffi.rs             # C FFI for Go/C#/Node.js (cdylib + staticlib)
├── crates/
│   ├── office_oxide_cli/  # CLI binary: office-oxide
│   └── office_oxide_mcp/  # MCP server binary: office-oxide-mcp
├── examples/
│   ├── rust/              # extract.rs, make_smoke.rs
│   ├── python/            # extract.py, read_xlsx.py, replace.py
│   ├── go/                # extract, read_xlsx, replace
│   ├── javascript/        # extract.mjs, read_xlsx.mjs, replace.mjs
│   └── c/                 # extract.c
├── python/                # Python package: office_oxide/__init__.py, _native.pyi
├── go/                    # Go bindings (CGo over C FFI)
├── js/                    # Node.js native bindings (koffi)
├── wasm-pkg/              # WASM npm package config
├── csharp/                # C# / .NET bindings (P/Invoke)
├── include/               # C header: office_oxide_c/office_oxide.h
└── docs/                  # Architecture, per-language getting-started guides
```

## Development Workflow

### 1. Pick a Task

- Check [Issues](https://github.com/yfedoseev/office_oxide/issues)
- Look for issues labeled `help-wanted` or `good-first-issue`
- Comment on the issue to claim it

### 2. Create a Branch

```bash
git checkout -b feature/your-feature-name
# or
git checkout -b fix/your-bug-fix
```

Branch naming:
- `feature/` - New features
- `fix/` - Bug fixes
- `docs/` - Documentation updates
- `test/` - Test additions
- `refactor/` - Code refactoring

### 3. Make Changes

Write code following our [Coding Standards](#coding-standards).

### 4. Test Your Changes

```bash
# Run all tests
cargo test

# Run tests for a specific module (all in the same crate)
cargo test docx::

# Run with features
cargo test --features python

# Build and verify examples
cargo build --examples

# Watch mode (auto-reload)
cargo watch -x test
```

### 5. Format and Lint

```bash
# Format code
cargo fmt

# Run linter
cargo clippy -- -D warnings

# Or use the Makefile
make check-all
```

### 6. Commit Your Changes

Follow [Conventional Commits](https://www.conventionalcommits.org/):

```bash
git commit -m "feat: add PDF object parser"
git commit -m "fix: correct unicode mapping in ToUnicode CMap"
git commit -m "docs: update API documentation"
```

Commit types:
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation only
- `test`: Adding tests
- `refactor`: Code refactoring
- `perf`: Performance improvement
- `chore`: Maintenance tasks

### 7. Push and Create Pull Request

```bash
git push origin feature/your-feature-name
```

Then create a pull request on GitHub.

## Coding Standards

### Rust

#### Style

- Follow [Rust API Guidelines](https://rust-lang.github.io/api-guidelines/)
- Use `rustfmt` (configured in `rustfmt.toml`)
- Maximum line length: 100 characters
- Use 4 spaces for indentation

#### Naming Conventions

```rust
// Modules and crates
mod document_parser;

// Types (structs, enums, traits)
struct DocxDocument;
enum DocumentFormat;
trait Extractable;

// Functions and methods
fn parse_document() -> Result<Document>;

// Constants
const MAX_RECURSION_DEPTH: usize = 100;
```

#### Error Handling

```rust
// Use Result<T> for fallible operations
pub fn open(path: &Path) -> Result<Document> {
    let file = File::open(path)?;
    let doc = parse_file(file)?;
    Ok(doc)
}

// Avoid unwrap() in library code (only in tests and examples)
```

#### Safety

- Avoid `unsafe` unless absolutely necessary
- Document all `unsafe` blocks with safety invariants
- Prefer safe abstractions from the standard library

### Python

- Follow [PEP 8](https://pep8.org/)
- Use `ruff` for formatting and linting
- Type hints for all public functions
- Docstrings in Google style

## Testing

### Unit Tests

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_document() {
        let doc = Document::open("tests/fixtures/simple.docx").unwrap();
        let text = doc.plain_text();
        assert!(text.contains("Hello"));
    }
}
```

### Integration Tests

Located in each crate's `tests/` directory.

### Coverage Goals

- **Library code**: 85%+ coverage (enforced in CI)
- **Critical paths**: 100% coverage (parsing, error handling)

Check coverage:
```bash
cargo llvm-cov --lib --tests --html
open target/llvm-cov/html/index.html
```

## Documentation

### Code Documentation

- All public items must have doc comments
- Include examples in doc comments
- Run `cargo doc --no-deps` to check rendered docs

### Examples

```rust
use office_oxide::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let doc = Document::open("report.docx")?;
    println!("{}", doc.plain_text());
    Ok(())
}
```

## Submitting Changes

### Pull Request Checklist

Before submitting a PR, ensure:

- [ ] Code compiles without warnings
- [ ] All tests pass (`cargo test`)
- [ ] Code is formatted (`cargo fmt`)
- [ ] Clippy passes (`cargo clippy -- -D warnings`)
- [ ] New code has tests
- [ ] Documentation is updated
- [ ] Commit messages follow conventions
- [ ] PR description explains changes clearly

### Review Process

1. Maintainers will review your PR
2. Address feedback and push updates
3. Once approved, your PR will be merged
4. Your changes will appear in the next release

## Developer Certificate of Origin (DCO)

All commits must carry a `Signed-off-by` trailer certifying that you wrote the code and have the right to contribute it under the project's MIT OR Apache-2.0 license. Add it with:

```bash
git commit -s -m "feat(docx): add heading extraction"
# Produces: Signed-off-by: Your Name <you@example.com>
```

This is checked automatically on every pull request by the DCO CI job. There is no CLA to sign — the sign-off is all that's required.

## License

By contributing, you agree that your contributions will be dual licensed under **MIT OR Apache-2.0**, without any additional terms or conditions.

This means:
- Your code will be available under permissive open-source licenses
- Users can choose either MIT or Apache-2.0 for their needs
- See [LICENSE-MIT](LICENSE-MIT) and [LICENSE-APACHE](LICENSE-APACHE) for full terms

## Questions?

- Read code comments and documentation
- Check `docs/` for architecture and specification docs
- Open an issue for questions
- Join discussions on GitHub Discussions

Thank you for contributing!
