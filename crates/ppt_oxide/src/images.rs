//! Image extraction from PPT files.
//!
//! Re-exports the shared BLIP parser from `cfb_oxide` with PPT-specific type aliases.

pub use cfb_oxide::blip::BlipFormat as ImageFormat;
pub use cfb_oxide::blip::BlipImage as PptImage;
pub use cfb_oxide::blip::extract_blip_images as extract_images;
