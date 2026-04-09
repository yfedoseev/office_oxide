//! # office_oxide::core
//!
//! Shared primitives for OOXML document processing.
//!
//! Provides the Open Packaging Conventions (OPC) layer shared by
//! DOCX, XLSX, and PPTX formats: ZIP archive handling, content types,
//! relationships, core properties, and DrawingML shared types.

pub mod content_types;
pub mod editable;
pub mod error;
pub mod opc;
pub mod parallel;
pub mod properties;
pub mod relationships;
pub mod theme;
pub mod traits;
pub mod units;
pub mod xml;

pub use error::{Error, Result};
pub use traits::OfficeDocument;
