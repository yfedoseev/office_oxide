# Makefile for office_oxide Python bindings
#
# Common development tasks for building and testing the Python package

.PHONY: dev install test build clean help lint-py fmt-py fmt-py-check check-py

# Development install (editable mode)
# Builds the Rust extension and installs the Python package in development mode
dev:
	maturin develop --features python

# Install in release mode
# Builds the optimized Rust extension and installs the Python package
install:
	maturin develop --release --features python

# Run Python tests
# Executes pytest on the tests/ directory
test:
	pytest tests/

# Run Python tests with verbose output
test-verbose:
	pytest tests/ -v

# Run Python tests with coverage
test-coverage:
	pytest tests/ --cov=office_oxide --cov-report=html

# Build wheel package
# Creates a distributable Python wheel in target/wheels/
build:
	maturin build --release --features python

# Build wheel for all Python versions
build-all:
	maturin build --release --features python --interpreter python3.8 python3.9 python3.10 python3.11 python3.12

# Clean build artifacts
# Removes all build artifacts and compiled extensions
clean:
	cargo clean
	rm -rf target/
	rm -rf python/office_oxide/*.so
	rm -rf python/office_oxide/*.pyd
	rm -rf python/office_oxide/__pycache__
	rm -rf tests/__pycache__
	rm -rf .pytest_cache
	rm -rf htmlcov/
	rm -rf .coverage

# Run Rust tests
test-rust:
	cargo test

# Run Rust tests with Python feature enabled
test-rust-python:
	cargo test --features python

# Run Clippy linter on Rust code
lint:
	cargo clippy --all-targets -- -D warnings

# Run Clippy linter with Python feature
lint-python:
	cargo clippy --features python --all-targets -- -D warnings

# Run Ruff linter on Python code
lint-py:
	ruff check .

# Auto-fix Python linting issues
lint-py-fix:
	ruff check --fix .

# Format Rust code
fmt:
	cargo fmt

# Format Python code with Ruff
fmt-py:
	ruff format .

# Check formatting without modifying files
fmt-check:
	cargo fmt -- --check

# Check Python formatting without modifying files
fmt-py-check:
	ruff format --check .

# Run all Rust checks (format, lint, test)
check: fmt-check lint test-rust

# Run all Python checks (format, lint)
check-py: fmt-py-check lint-py

# Run all checks for both Rust and Python
check-all: check check-py

# Display help
help:
	@echo "office_oxide - Makefile Commands"
	@echo ""
	@echo "Development:"
	@echo "  make dev              - Install in development mode (fast rebuilds)"
	@echo "  make install          - Install in release mode (optimized)"
	@echo ""
	@echo "Testing:"
	@echo "  make test             - Run Python tests"
	@echo "  make test-verbose     - Run Python tests with verbose output"
	@echo "  make test-coverage    - Run Python tests with coverage report"
	@echo "  make test-rust        - Run Rust tests"
	@echo "  make test-rust-python - Run Rust tests with Python feature"
	@echo ""
	@echo "Building:"
	@echo "  make build            - Build release wheel"
	@echo "  make build-all        - Build wheels for all Python versions"
	@echo ""
	@echo "Code Quality (Rust):"
	@echo "  make lint             - Run Clippy linter on Rust code"
	@echo "  make lint-python      - Run Clippy linter with Python feature"
	@echo "  make fmt              - Format Rust code"
	@echo "  make fmt-check        - Check Rust formatting without modifying"
	@echo "  make check            - Run all Rust checks (format, lint, test)"
	@echo ""
	@echo "Code Quality (Python):"
	@echo "  make lint-py          - Run Ruff linter on Python code"
	@echo "  make lint-py-fix      - Auto-fix Python linting issues"
	@echo "  make fmt-py           - Format Python code with Ruff"
	@echo "  make fmt-py-check     - Check Python formatting without modifying"
	@echo "  make check-py         - Run all Python checks (format, lint)"
	@echo ""
	@echo "Code Quality (All):"
	@echo "  make check-all        - Run all checks for both Rust and Python"
	@echo ""
	@echo "Cleanup:"
	@echo "  make clean            - Remove all build artifacts"
	@echo ""
	@echo "Help:"
	@echo "  make help             - Display this help message"
