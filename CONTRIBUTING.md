# Contributing to html2pptx

Thank you for your interest in contributing! Here are some guidelines to help you get started.

## Development Setup

```bash
git clone https://github.com/junyuaini/html2pptx.git
cd html2pptx
npm install
```

## Project Structure

```
src/
├── cli.js                    # CLI entry point (Commander.js)
├── index.js                  # Main API — convert() function
├── extract.js                # Puppeteer element extraction
├── generate.js               # PPTX generation (PptxGenJS)
├── gradient-postprocess.js   # OOXML gradient injection
└── utils.js                  # Shared utilities (color, font, coordinates)
```

## Architecture

The conversion pipeline follows a three-layer architecture:

1. **Input Layer** — CLI or API receives HTML file(s)
2. **Extraction Layer** (`extract.js`) — Puppeteer renders HTML, extracts element positions and computed styles via `getBoundingClientRect()` and `getComputedStyle()`
3. **Generation Layer** (`generate.js`) — PptxGenJS creates native PowerPoint shapes from extracted data
4. **Post-processing** (`gradient-postprocess.js`) — Unzips PPTX, replaces solid fills with native OOXML gradients

## Code Style

- Keep functions under 100 lines
- Keep files under 1000 lines
- Maximum 3 levels of nesting
- No magic values — use named constants
- Follow the existing pattern: input validation → business validation → business logic
- Add error handling with clear, specific messages
- Add logging at key business nodes

## Testing

```bash
# Run the built-in test
npm test

# Test with custom files
node src/cli.js test/cover.html -o test/output.pptx -v
```

Place test HTML files in the `test/` directory.

## Pull Request Process

1. Fork the repository and create a feature branch
2. Make your changes following the code style guidelines
3. Test your changes with the built-in test command
4. Update documentation if needed
5. Submit a pull request with a clear description of changes

## Reporting Issues

When reporting bugs, please include:

- Node.js version (`node -v`)
- Operating system
- A minimal HTML file that reproduces the issue
- Expected vs actual behavior
- Error logs (if any)

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
