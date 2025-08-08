# Markdown-DOCX Converter

🚀 **NodeJS Markdown ↔ DOCX converter** with advanced features including Mermaid diagram support, professional styling templates, and comprehensive link processing.

[![Build Status](https://img.shields.io/badge/build-passing-green.svg)](/)
[![Coverage](https://img.shields.io/badge/coverage-85%25-green.svg)](/)
[![Version](https://img.shields.io/badge/version-0.5.0-blue.svg)](/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](/)

---

## 🎯 **Overview**

A Node.js application designed to convert Markdown files (with Mermaid support) to DOCX format and vice versa Perfect for technical documentation, reports, and automated document workflows.

---

## ✨ **Features**

###  **Core Functionality**
- **Bidirectional conversion** between Markdown and DOCX formats
- **Professional document templates** for different use cases
- **Mermaid diagram processing** with automatic image generation
- **Advanced link handling** (internal anchors, external URLs, references)
- **Table of contents generation** with navigable links
- **Batch processing** for multiple files

###  **Professional Styling**
- **7 built-in templates**: Professional Report, Technical Documentation, Business Proposal, Academic Paper, Modern, Classic, Simple
- **Customizable styling** for headings, paragraphs, code blocks, tables
- **Professional typography** with proper font selection and spacing
- **Page layout control** (margins, orientation, headers/footers)

###  **Advanced Link Processing**
- **Internal link preservation** with automatic anchor generation
- **External link validation** and formatting
- **Reference-style links** support
- **Cross-reference generation** for headings and sections

###  **Mermaid Diagram Support**
- **Multiple diagram types**: Flowcharts, Sequence, Class, State, ER, Gantt, etc.
- **5 built-in themes**: Default, Forest, Dark, Neutral, Base
- **High-quality image generation** (SVG/PNG)
- **Automatic diagram validation** and error handling

###  **Developer Experience**
- **TypeScript support** with full type definitions
- **Comprehensive error handling** with detailed messages
- **Performance monitoring** and optimization
- **Extensive logging** for debugging
- **Memory-efficient processing** for large documents


##  Installation

```bash
npm install markdown-docx-converter
```

For global CLI usage:

```bash
npm install -g markdown-docx-converter
```

##  CLI (Command Line) Usage

You can run the CLI globally (md-docx), locally (node dist/cli/index.js), or via npm script (npm run cli -- …).

- Global: md-docx <command> [options]
- Local: node dist/cli/index.js <command> [options]
- NPM: npm run cli -- <command> [options]

Main commands:
- convert <input.md>  → DOCX
- extract <input.docx> → Markdown
- batch <input-dir>   → bulk convert
- validate <input.md> → check file
- stats <input>        → show stats
- list                 → templates & themes

Common options:
- -o, --output <file>
- -t, --template <name> (default: professional-report)
- -m, --mermaid-theme <theme> (default: default)
- --title <text> --author <text> --subject <text>
- --toc (table of contents)
- --no-links (disable link processing)
- --orientation <portrait|landscape>
- --verbose

Examples (PowerShell):

```powershell
# Convert Markdown → DOCX with TOC (Make sure that an output folder exists or adapt the -o parameter)
node dist/cli/index.js convert .\demo-complete.md -o .\output\demo-complete.docx --template professional-report --toc

# Global CLI (if installed with -g)
md-docx convert .\sample.md -o .\output\sample.docx --toc --title "Demo" --author "Team"

# Extract DOCX → Markdown with images
node dist/cli/index.js extract .\sample.docx -o .\extracted.md --extract-images --image-dir images --image-format png

# Batch convert a folder to DOCX
node dist/cli/index.js batch .\docs -o .\output -f docx --template professional-report --toc

# Validate and show stats
node dist/cli/index.js validate .\demo-complete.md ; node dist/cli/index.js stats .\demo-complete.md
```


##  Quick Start

### Programmatic Usage

```typescript
import { MarkdownDocxConverter } from 'markdown-docx-converter';

const converter = new MarkdownDocxConverter({
  template: 'professional-report',
  mermaidTheme: 'forest',
  preserveLinks: true,
  tocGeneration: true
});

// Convert Markdown to DOCX
const docxResult = await converter.markdownToDocx(markdownContent, {
  title: 'Technical Report',
  author: 'John Doe',
  includeMetadata: true
});

if (docxResult.success) {
  await fs.writeFile('report.docx', docxResult.output);
  console.log('DOCX created successfully!');
}

// Convert DOCX to Markdown
const markdownResult = await converter.docxToMarkdown('document.docx', {
  preserveFormatting: true,
  extractImages: true,
  imageOutputDir: './images'
});

if (markdownResult.success) {
  await fs.writeFile('document.md', markdownResult.output);
  console.log('Markdown extracted successfully!');
}
```

### CLI Usage

```bash
# Convert Markdown to DOCX
md-docx convert input.md -o output.docx --template professional-report --toc

# Extract DOCX to Markdown
md-docx extract document.docx -o output.md --extract-images --image-dir images

# Batch convert all Markdown files in a directory
md-docx batch ./docs -o ./output --format docx --template technical-documentation

# Validate Markdown file
md-docx validate document.md

# Show file statistics
md-docx stats document.md

# List available templates and themes
md-docx list
```

##  Documentation

### Available Templates

| Template | Description | Best For |
|----------|-------------|----------|
| `professional-report` | Clean, corporate styling with blue accents | Business reports, project documentation |
| `technical-documentation` | Developer-friendly with code highlighting | API docs, technical guides |
| `business-proposal` | Elegant typography with traditional styling | Proposals, contracts, formal documents |
| `academic-paper` | IEEE/ACM style formatting | Research papers, academic writing |
| `modern` | Contemporary design with minimal styling | Presentations, modern documentation |
| `classic` | Traditional book-style formatting | Traditional documents, literature |
| `simple` | Clean, minimal formatting | General purpose, simple documents |

### Mermaid Diagram Support

The converter supports all Mermaid diagram types:

```markdown
# Flowchart Example
\`\`\`mermaid
graph TD
    A[Start] --> B{Decision}
    B -->|Yes| C[Process]
    B -->|No| D[Alternative]
    C --> E[End]
    D --> E
\`\`\`

# Sequence Diagram
\`\`\`mermaid
sequenceDiagram
    participant User
    participant API
    participant Database
    
    User->>API: Request
    API->>Database: Query
    Database-->>API: Result
    API-->>User: Response
\`\`\`
```

### Advanced Configuration

```typescript
import { MarkdownDocxConverter } from 'markdown-docx-converter';

const converter = new MarkdownDocxConverter();

await converter.markdownToDocx(content, {
  // Document metadata
  title: 'My Document',
  author: 'Author Name',
  subject: 'Document Subject',
  keywords: ['keyword1', 'keyword2'],
  
  // Styling options
  template: 'professional-report',
  orientation: 'portrait', // or 'landscape'
  margins: {
    top: 1440,    // 1 inch in twips
    right: 1440,
    bottom: 1440,
    left: 1440
  },
  
  // Feature toggles
  tocGeneration: true,
  preserveLinks: true,
  includeMetadata: true,
  
  // Mermaid options
  mermaidTheme: 'forest',
  
  // Custom styles
  customStyles: {
    headings: {
      h1: {
        fontSize: 32,
        color: '2F5597',
        bold: true
      }
    },
    paragraph: {
      fontSize: 24,
      alignment: 'justify'
    }
  }
});
```

### Error Handling

```typescript
const result = await converter.markdownToDocx(content);

if (!result.success) {
  console.error('Conversion failed:', result.error?.message);
  console.error('Error code:', result.error?.code);
  console.error('Details:', result.error?.details);
} else {
  console.log('Success! Metadata:', result.metadata);
  
  if (result.warnings?.length > 0) {
    console.warn('Warnings:', result.warnings);
  }
}
```

## 🏗️ **Architecture**

The converter is built with a clean, modular architecture designed for maintainability and extensibility:

```
markdown-docx-converter/
├── 📁 src/                          # Source code
│   ├── 📁 converters/               # Core conversion engines
│   │   ├── 📄 markdown-to-docx.ts   # MD → DOCX conversion logic
│   │   └── 📄 docx-to-markdown.ts   # DOCX → MD extraction logic
│   ├── 📁 processors/               # Specialized content processors
│   │   ├── 📄 mermaid-processor.ts  # Diagram rendering & validation
│   │   └── 📄 link-processor.ts     # Link handling & TOC generation
│   ├── 📁 styles/                   # Document styling & templates
│   │   └── 📄 document-templates.ts # Professional template definitions
│   ├── 📁 utils/                    # Shared utilities & helpers
│   │   ├── 📄 logger.ts             # Logging system
│   │   └── 📄 index.ts              # Utility exports
│   ├── 📁 types/                    # TypeScript type definitions
│   │   └── 📄 index.ts              # Type exports & interfaces
│   ├── 📁 cli/                      # Command-line interface
│   │   └── 📄 index.ts              # CLI commands & argument parsing
│   └── 📄 index.ts                  # Main library entry point
├── 📁 tests/                        # Comprehensive test suites
│   ├── 📄 integration.test.ts       # End-to-end conversion tests
│   └── 📄 unit.test.ts              # Component-level tests
├── 📁 templates/                    # Document template assets
│   ├── 📄 professional-report.json # Business report styling
│   └── 📄 technical-docs.json      # Technical documentation styling
└── 📁 docs/                         # Project documentation
    ├── 📄 API.md                    # API reference
    └── 📄 CONTRIBUTING.md           # Developer guidelines
```

### **Key Components**

#### **MarkdownToDocxConverter**
- Parses Markdown using `marked`
- Processes Mermaid diagrams with Puppeteer
- Generates professional DOCX with `docx` library
- Handles links, tables, code blocks, and formatting

#### **DocxToMarkdownConverter**
- Extracts content using `mammoth`
- Converts HTML to Markdown with `turndown`
- Preserves formatting and extracts images
- Handles complex document structures

#### **MermaidProcessor**
- Validates Mermaid syntax
- Renders diagrams to SVG/PNG
- Supports multiple themes
- Optimizes image quality and size

#### **LinkProcessor**
- Processes internal and external links
- Generates table of contents
- Creates Word bookmarks and hyperlinks
- Validates link targets

##  Testing

Run the test suite:

```bash
npm test
```

Run with coverage:

```bash
npm run test:coverage
```

### Test Categories

- **Unit Tests**: Individual component testing
- **Integration Tests**: End-to-end conversion testing
- **Performance Tests**: Large document processing
- **Error Handling**: Edge cases and error scenarios

##  Performance

### Benchmarks

| Document Size | Processing Time | Memory Usage |
|---------------|----------------|--------------|
| Small (< 1MB) | < 2 seconds | < 100MB |
| Medium (1-5MB) | < 10 seconds | < 200MB |
| Large (5-10MB) | < 30 seconds | < 500MB |

### Performance Tips

1. **Use appropriate templates** - simpler templates process faster
2. **Optimize Mermaid diagrams** - reduce complexity for better performance
3. **Batch processing** - more efficient for multiple files
4. **Memory management** - process large files in chunks

---

## � Known Bugs and Limitations

These are the current known constraints of the converter. If you hit one of them, please open an issue with a minimal sample DOCX/Markdown.

- DOCX → Markdown: Complex tables (heavy rowspans/colspans or nested tables) are approximated as GitHub-Flavored Markdown tables. A reflow heuristic may adjust layout when Mammoth produces one cell per row. Manual review is recommended for complex tables.
- DOCX → Markdown fallback: In rare edge cases, Mammoth’s HTML conversion can fail. The converter will fall back to plain text extraction, which loses formatting, images, and links. Workarounds: re-save the DOCX in Word/LibreOffice, remove SmartArt/text boxes, or simplify the problematic section and retry.
- Internal links and bookmarks: Word bookmarks aren’t always exported. We inject anchors for headings, but cross-references to arbitrary bookmarks may not resolve automatically. You may need to fix those links manually in the extracted Markdown.
- Code blocks: Some pre/code regions can become multiple adjacent fenced blocks. Merge them manually if needed.
- Images: Extracted image paths in Markdown use POSIX-style separators (/) for cross-platform compatibility. Some embedded/auto-generated Word drawings (SmartArt, shapes) may not be exported by Mammoth.
- Unsupported DOCX features on extraction: tracked changes, comments, headers/footers content, footnotes/endnotes, text boxes, and floating shapes are not preserved and may be dropped or converted as plain text.
- Markdown → DOCX image sizing: Mermaid diagrams are sized using intrinsic bitmap metadata. Very large diagrams may still be scaled by Word. If sizing looks off, adjust the Mermaid configuration or set a maximum width.

We’re iterating on these—feedback and sample files help prioritize fixes.

---

## �📚 **Dependencies**

This project is built on top of several excellent open-source libraries:

### **Core Dependencies**

| Package | Version | Purpose | Author | Repository |
|---------|---------|---------|--------|------------|
| **[docx](https://www.npmjs.com/package/docx)** | ^8.5.0 | DOCX document generation | [Dolan Miu](https://github.com/dolanmiu) | [dolanmiu/docx](https://github.com/dolanmiu/docx) |
| **[marked](https://www.npmjs.com/package/marked)** | ^9.1.6 | Markdown parser and compiler | [Christopher Jeffrey](https://github.com/chjj) | [markedjs/marked](https://github.com/markedjs/marked) |
| **[mermaid](https://www.npmjs.com/package/mermaid)** | ^10.6.1 | Diagram and flowchart generation | [Mermaid Team](https://github.com/mermaid-js) | [mermaid-js/mermaid](https://github.com/mermaid-js/mermaid) |
| **[puppeteer](https://www.npmjs.com/package/puppeteer)** | ^21.5.2 | Headless Chrome automation | [Google Chrome Team](https://github.com/puppeteer) | [puppeteer/puppeteer](https://github.com/puppeteer/puppeteer) |
| **[sharp](https://www.npmjs.com/package/sharp)** | ^0.34.3 | High-performance image processing | [Lovell Fuller](https://github.com/lovell) | [lovell/sharp](https://github.com/lovell/sharp) |
| **[commander](https://www.npmjs.com/package/commander)** | ^11.1.0 | Command-line interface framework | [TJ Holowaychuk](https://github.com/tj) | [tj/commander.js](https://github.com/tj/commander.js) |
| **[fs-extra](https://www.npmjs.com/package/fs-extra)** | ^11.1.1 | Extended filesystem operations | [JP Richardson](https://github.com/jprichardson) | [jprichardson/node-fs-extra](https://github.com/jprichardson/node-fs-extra) |
| **[canvas](https://www.npmjs.com/package/canvas)** | ^3.1.2 | HTML5 Canvas API for Node.js | [Automattic](https://github.com/Automattic) | [Automattic/node-canvas](https://github.com/Automattic/node-canvas) |
| **[mammoth](https://www.npmjs.com/package/mammoth)** | ^1.6.0 | DOCX to HTML/Markdown conversion | [Mike Williams](https://github.com/mwilliamson) | [mwilliamson/python-mammoth](https://github.com/mwilliamson/python-mammoth) |
| **[turndown](https://www.npmjs.com/package/turndown)** | ^7.1.2 | HTML to Markdown converter | [Dom Christie](https://github.com/domchristie) | [mixmark-io/turndown](https://github.com/mixmark-io/turndown) |
| **[chalk](https://www.npmjs.com/package/chalk)** | ^4.1.2 | Terminal string styling | [Sindre Sorhus](https://github.com/sindresorhus) | [chalk/chalk](https://github.com/chalk/chalk) |
| **[ora](https://www.npmjs.com/package/ora)** | ^5.4.1 | Elegant terminal spinners | [Sindre Sorhus](https://github.com/sindresorhus) | [sindresorhus/ora](https://github.com/sindresorhus/ora) |
| **[uuid](https://www.npmjs.com/package/uuid)** | ^9.0.1 | RFC4122 UUID generator | [Robert Kieffer](https://github.com/broofa) | [uuidjs/uuid](https://github.com/uuidjs/uuid) |
| **[mime-types](https://www.npmjs.com/package/mime-types)** | ^2.1.35 | MIME type utilities | [Jonathan Ong](https://github.com/jongleberry) | [jshttp/mime-types](https://github.com/jshttp/mime-types) |
| **[ajv](https://www.npmjs.com/package/ajv)** | ^8.12.0 | JSON Schema validator | [Evgeny Poberezkin](https://github.com/epoberezkin) | [ajv-validator/ajv](https://github.com/ajv-validator/ajv) |

### **Development Dependencies**

| Package | Purpose | Author | Repository |
|---------|---------|--------|------------|
| **[@types/node](https://www.npmjs.com/package/@types/node)** | TypeScript definitions for Node.js | [DefinitelyTyped](https://github.com/DefinitelyTyped) | [DefinitelyTyped/DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped) |
| **[typescript](https://www.npmjs.com/package/typescript)** | TypeScript compiler | [Microsoft](https://github.com/microsoft) | [microsoft/TypeScript](https://github.com/microsoft/TypeScript) |
| **[jest](https://www.npmjs.com/package/jest)** | JavaScript testing framework | [Meta](https://github.com/facebook) | [facebook/jest](https://github.com/facebook/jest) |
| **[ts-jest](https://www.npmjs.com/package/ts-jest)** | TypeScript preprocessor for Jest | [Kulshekhar Kabra](https://github.com/kulshekhar) | [kulshekhar/ts-jest](https://github.com/kulshekhar/ts-jest) |
| **[@types/jest](https://www.npmjs.com/package/@types/jest)** | TypeScript definitions for Jest | [DefinitelyTyped](https://github.com/DefinitelyTyped) | [DefinitelyTyped/DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped) |
| **[eslint](https://www.npmjs.com/package/eslint)** | JavaScript and TypeScript linting | [Nicholas C. Zakas](https://github.com/nzakas) | [eslint/eslint](https://github.com/eslint/eslint) |
| **[@typescript-eslint/eslint-plugin](https://www.npmjs.com/package/@typescript-eslint/eslint-plugin)** | TypeScript ESLint rules | [TypeScript ESLint](https://github.com/typescript-eslint) | [typescript-eslint/typescript-eslint](https://github.com/typescript-eslint/typescript-eslint) |
| **[@typescript-eslint/parser](https://www.npmjs.com/package/@typescript-eslint/parser)** | TypeScript parser for ESLint | [TypeScript ESLint](https://github.com/typescript-eslint) | [typescript-eslint/typescript-eslint](https://github.com/typescript-eslint/typescript-eslint) |

### **Special Thanks**

We extend our gratitude to all the maintainers and contributors of these fantastic open-source projects that make this converter possible:

- 🙏 **Dolan Miu** for the excellent `docx` library that powers our DOCX generation
- 🙏 **Mermaid Team** for the amazing diagram generation capabilities
- 🙏 **Google Chrome Team** for Puppeteer's reliable automation
- 🙏 **Christopher Jeffrey** and the Marked.js team for robust Markdown parsing
- 🙏 **Lovell Fuller** for Sharp's lightning-fast image processing
- 🙏 **Mike Williams** for Mammoth's excellent DOCX parsing
- 🙏 **Dom Christie** for Turndown's HTML to Markdown conversion
- 🙏 **Sindre Sorhus** for the beautiful CLI experience with Chalk and Ora

---

## 🛠️ **Development**

### Prerequisites

- Node.js 18+ (LTS recommended)
- TypeScript 5.0+
- Git

### Setup

```bash
# Clone the repository
git clone <repository-url>
cd markdown-docx-converter

# Install dependencies
npm install

# Build the project
npm run build

# Run tests
npm test

# Start development mode
npm run dev
```

### Building

```bash
npm run build
```

### Linting

```bash
npm run lint
npm run lint:fix
```

##  Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

### Development Workflow

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass: `npm test`
6. Commit your changes: `git commit -m 'Add amazing feature'`
7. Push to the branch: `git push origin feature/amazing-feature`
8. Open a Pull Request

##  License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

##  Support

- **Documentation**: Check this README and inline code documentation
- **Issues**: Report bugs and feature requests on [GitHub Issues](https://github.com/your-repo/issues)
- **Discussions**: Join the community on [GitHub Discussions](https://github.com/your-repo/discussions)

##  Roadmap

### Version 1.0
- FIXME : 
  - [ ] Conversion of docx to mark down has issues, some are known and some are still to be discoverd :-)
    - [ ] Nested lists conversion not working properly
    - [ ] Internal link conversion not working properly
  - [ ] CLI arguments should have default values so that it can be used without specifying all of them each time

### Version 1.1 (Future)
- [ ] Additional document templates
- [ ] Enhanced Mermaid theme customization
- [ ] PDF export support

### Version 1.2 (Future)
- [ ] Advanced image processing
- [ ] Plugin system for custom processors

##  Changelog


---

<div align="center">

**Made with ❤️ for the documentation community**

[⭐ Star this project](https://github.com/ndebras/md-docx-converter) • [🐛 Report Bug](https://github.com/ndebras/md-docx-converter/issues) • [💡 Request Feature](https://github.com/ndebras/md-docx-converter/issues)

</div>
