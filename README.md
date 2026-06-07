# html2pptx

[![npm version](https://img.shields.io/npm/v/html2pptx)](https://www.npmjs.com/package/html2pptx)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js Version](https://img.shields.io/badge/node-%3E%3D18.0.0-brightgreen)](https://nodejs.org)

> Convert static HTML/CSS pages to editable PowerPoint (.pptx) presentations.

**html2pptx** renders your HTML pages in a headless browser, extracts every visible element's position and style, then rebuilds them as native PowerPoint shapes — text boxes, rectangles, tables, images, and more. The output is a fully editable `.pptx` file.

---

## ✨ Features

- **Text Elements** — h1–h6, p, span, a, li with full font styling
- **Rich Text** — mixed bold / italic / underline / color within a single element
- **Containers** — div, section, header, footer with background-color, border, border-radius
- **Tables** — table, tr, td, th with cell-level styles and row heights
- **Images & SVG** — img elements and inline SVG converted to embedded images
- **CSS Gradients** — linear-gradient with multiple color stops, angles, and transparency → native OOXML gradFill
- **Pseudo-elements** — ::before and ::after with position, background, border, and text content
- **CSS Transforms** — rotate, translate, scale
- **Batch Processing** — single file, multiple files, or entire directory

---

## 📦 Installation

```bash
npm install -g html2pptx
```

> **Requirements:** Node.js >= 18.0.0. Puppeteer will automatically download a compatible Chromium binary on first install.

---

## 🚀 Quick Start

### CLI

```bash
# Single HTML file
html2pptx slide.html -o output.pptx

# Multiple HTML files → one PPT (one slide per file)
html2pptx slide1.html slide2.html slide3.html -o output.pptx

# All HTML files in a directory
html2pptx ./slides/ -o output.pptx

# Verbose logging
html2pptx slide.html -o output.pptx --verbose
```

### API

```javascript
const { convert } = require('html2pptx');

await convert('input.html', 'output.pptx', {
  scale: 0.75,    // px → pt scale factor (default: 0.75)
  verbose: true,  // enable detailed logs
});
```

---

## 📐 How It Works

```
HTML File(s)
    │
    ▼
┌─────────────────────────────┐
│  1. Puppeteer Rendering     │  ← Headless Chrome, 1280×720 viewport
│     getBoundingClientRect() │
│     getComputedStyle()      │
└─────────────┬───────────────┘
              │
              ▼
┌─────────────────────────────┐
│  2. Element Extraction      │  ← extract.js
│     • Text runs (rich text) │
│     • Container shapes      │
│     • Tables & cells        │
│     • Images & SVG          │
│     • Pseudo-elements       │
│     • CSS transforms        │
└─────────────┬───────────────┘
              │
              ▼
┌─────────────────────────────┐
│  3. PPTX Generation         │  ← generate.js (PptxGenJS)
│     • Text boxes            │
│     • Rectangles / ovals    │
│     • Tables                │
│     • Images                │
└─────────────┬───────────────┘
              │
              ▼
┌─────────────────────────────┐
│  4. Gradient Post-process   │  ← gradient-postprocess.js
│     solidFill → gradFill    │     (native OOXML gradients)
└─────────────┬───────────────┘
              │
              ▼
         output.pptx
```

---

## 🎨 Supported HTML/CSS

| Category | Supported |
|----------|-----------|
| **Text** | h1–h6, p, span, a, li, b, strong, i, em, u, br |
| **Containers** | div, section, article, header, footer, main, nav, aside |
| **Tables** | table, tr, td, th, thead, tbody, caption |
| **Media** | img, svg |
| **Decorations** | hr, ::before, ::after |
| **Background** | background-color, linear-gradient |
| **Border** | border, border-radius, individual side borders |
| **Typography** | font-family, font-size, font-weight, font-style, color, text-align, line-height, text-decoration |
| **Layout** | padding, opacity, transform (rotate / translate / scale) |
| **Rich Text** | nested b / i / u / span with inherited styles |

---

## 📝 HTML Authoring Guide

For best conversion results, follow these guidelines when creating HTML pages:

- **Fixed size:** 1280 × 720 px (16:9)
- **No JavaScript** — pure static HTML + CSS only
- **No external CSS frameworks** (Tailwind, Bootstrap, etc.)
- **No box-shadow** — use borders or pseudo-elements instead
- **No animations / transitions**
- **No CSS variables** (`var(--xxx)`) — use concrete color values
- **Safe fonts:** Microsoft YaHei, Arial, SimHei, SimSun
- **Colors:** hex (`#RRGGBB`) or `rgb()` / `rgba()`
- **Images:** absolute URLs (`https://...`)

See [AI生成H5提示词.md](./AI生成H5提示词.md) for ready-to-use AI prompts that generate compliant HTML pages.

---

## 📂 Project Structure

```
html2pptx/
├── src/
│   ├── cli.js                    # CLI entry point
│   ├── index.js                  # Main API (convert function)
│   ├── extract.js                # Puppeteer element extraction
│   ├── generate.js               # PPTX generation (PptxGenJS)
│   ├── gradient-postprocess.js   # Native OOXML gradient injection
│   └── utils.js                  # Color, font, coordinate utilities
├── test/                         # Sample HTML files for testing
├── AI生成H5提示词.md              # AI prompt templates (Chinese)
├── package.json
├── LICENSE
├── CHANGELOG.md
└── README.md
```

---

## 🔧 CLI Options

| Option | Alias | Description | Default |
|--------|-------|-------------|---------|
| `<inputs...>` | — | HTML file(s) or directory (required) | — |
| `--output <path>` | `-o` | Output .pptx file path | `output.pptx` |
| `--scale <number>` | `-s` | px → pt scale factor | `0.75` |
| `--verbose` | — | Enable detailed logging | `false` |
| `--version` | `-V` | Show version number | — |

---

## 🤝 Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines.

---

## 📄 License

[MIT](./LICENSE) © 2026 junyuaini
