//! # office_oxide::core
//!
//! Shared primitives for OOXML document processing.
//!
//! Provides the Open Packaging Conventions (OPC) layer shared by
//! DOCX, XLSX, and PPTX formats: ZIP archive handling, content types,
//! relationships, core properties, and DrawingML shared types.

/// Standard base64, shared by the IR JSON image codec and the renderers.
pub(crate) mod base64;
/// Text extraction from DrawingML chart parts (`c:chartSpace`), shared by
/// the DOCX and PPTX readers for charts embedded via a drawing relationship.
pub mod chart;
pub(crate) mod codepage;
/// `[Content_Types].xml` parsing and writing.
pub mod content_types;
/// Shared `docProps/core.xml` generator used by DOCX, PPTX, XLSX writers.
pub mod core_properties;
/// MD5, SHA-1, SHA-2, AES decryption and RC4, for the default-password path.
#[cfg(feature = "default-password")]
pub(crate) mod crypto;
/// In-place editing of OPC packages (preserves unchanged parts).
pub mod editable;
pub(crate) mod embedded;
/// Helpers for embedding TrueType / OpenType font programs in DOCX,
/// PPTX, and XLSX packages.
pub mod embedded_fonts;
/// Core error type and `Result` alias used throughout OOXML parsing.
pub mod error;
/// Markdown escaping shared by every renderer.
pub mod markdown;
/// Decrypting a document protected with Office's default password.
#[cfg(feature = "default-password")]
pub(crate) mod offcrypto;
/// OPC (Open Packaging Conventions) reader and writer for ZIP-based packages.
pub mod opc;
/// Parallel processing helpers (Rayon-based, feature-gated).
pub mod parallel;
/// Core and extended document properties (`docProps/core.xml`, `docProps/app.xml`).
pub mod properties;
/// Relationship parsing, lookup, and serialization (`.rels` files).
pub mod relationships;

pub(crate) mod symbol_font;
/// DrawingML theme parsing: color schemes, font schemes, color resolution.
pub mod theme;
/// `OfficeDocument` trait implemented by all document types.
pub mod traits;
/// Unit types: `Emu`, `Twip`, `HalfPoint`, `Percentage1000`, `Angle60k`.
pub mod units;
/// The bounded record of the parts a reader could not use.
pub(crate) mod unreadable;
/// XML reading utilities, namespace constants, and attribute helpers.
pub mod xml;

pub use error::{Error, Result};
pub use traits::OfficeDocument;
