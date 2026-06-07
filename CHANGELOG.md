# Changelog

All notable changes to this project will be documented in this file.

## [1.1.0] - 2026-06-07

### Added
- Native OOXML gradient support via post-processing (linear-gradient with multi-stop, angle, transparency)
- Pseudo-element (::before / ::after) extraction and rendering
- CSS transform support (rotate, translate, scale)
- Rich text support with nested b/i/u/span style inheritance
- Table extraction with cell-level styles and row heights
- SVG element extraction (converted to Base64 embedded images)
- Individual side border support (border-top/right/bottom/left)
- Batch processing: directory mode and multi-file input

### Changed
- Improved shape minimum size threshold (0.01 inch) for fine decorative elements
- Enhanced font resolution with fallback chain

### Fixed
- Various edge cases in element positioning and style extraction

## [1.0.0] - 2026-06-05

### Added
- Initial release
- Core HTML-to-PPTX conversion pipeline
- Puppeteer-based element extraction (position, size, computed styles)
- PptxGenJS-based PPTX generation
- CLI interface with Commander.js
- Programmatic API (convert function)
- Support for text elements, containers, images, tables, hr
- Support for background-color, border, border-radius, font styles, padding, opacity
- 1280×720 H5 → 960×540pt PPT (16:9) with 0.75 scale factor
