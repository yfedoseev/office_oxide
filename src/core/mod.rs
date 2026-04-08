//! # office_oxide::core
//!
//! Shared primitives for OOXML document processing.
//!
//! Provides the Open Packaging Conventions (OPC) layer shared by
//! DOCX, XLSX, and PPTX formats: ZIP archive handling, content types,
//! relationships, core properties, and DrawingML shared types.

pub mod error;
pub mod units;
pub mod xml;
pub mod content_types;
pub mod relationships;
pub mod opc;
pub mod properties;
pub mod theme;
pub mod editable;
pub mod parallel;
pub mod traits;

pub use error::{Error, Result};
pub use traits::OfficeDocument;
