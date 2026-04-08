//! Image extraction from PPT files.
//!
//! Re-exports the shared BLIP parser from `cfb` with PPT-specific type aliases.

pub use crate::cfb::blip::BlipFormat as ImageFormat;
pub use crate::cfb::blip::BlipImage as PptImage;
pub use crate::cfb::blip::extract_blip_images as extract_images;
