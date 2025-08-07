import { MarkdownDocxConverter } from '../src';
import * as fs from 'fs-extra';
import * as path from 'path';

describe('MarkdownDocxConverter Integration Tests', () => {
  let converter: MarkdownDocxConverter;
  const testDataDir = path.join(__dirname, 'test-data');
  const outputDir = path.join(__dirname, 'output');

  beforeAll(async () => {
    converter = new MarkdownDocxConverter();
    await fs.ensureDir(testDataDir);
    await fs.ensureDir(outputDir);
  });

  afterAll(async () => {
    await fs.remove(outputDir);
  });

  describe('Markdown to DOCX Conversion', () => {
    test('should convert simple markdown to DOCX', async () => {
      const markdown = `# Test Document

This is a **bold** text and this is *italic*.

## Section 2

Here's a list:
- Item 1
- Item 2
- Item 3

\`\`\`javascript
function hello() {
  console.log('Hello, World!');
}
\`\`\`

| Column 1 | Column 2 |
|----------|----------|
| Data 1   | Data 2   |
`;

      const result = await converter.markdownToDocx(markdown, {
        title: 'Test Document',
        author: 'Test Author',
        template: 'professional-report',
      });

      expect(result.success).toBe(true);
      expect(result.output).toBeDefined();
      expect(result.metadata).toBeDefined();
      expect(result.metadata!.inputSize).toBeGreaterThan(0);
      expect(result.metadata!.outputSize).toBeGreaterThan(0);
    }, 15000);

    test('should convert markdown with Mermaid diagrams', async () => {
      const markdown = `# Diagram Test

Here's a flowchart:

\`\`\`mermaid
graph TD
    A[Start] --> B{Decision}
    B -->|Yes| C[Action 1]
    B -->|No| D[Action 2]
    C --> E[End]
    D --> E
\`\`\`

And here's a sequence diagram:

\`\`\`mermaid
sequenceDiagram
    participant A as Alice
    participant B as Bob
    A->>B: Hello Bob!
    B-->>A: Hi Alice!
\`\`\`
`;

      const result = await converter.markdownToDocx(markdown, {
        title: 'Mermaid Test',
        mermaidTheme: 'forest',
      });

      expect(result.success).toBe(true);
      expect(result.metadata!.mermaidDiagramCount).toBeGreaterThan(0);
    }, 20000);

    test('should handle internal links and generate TOC', async () => {
      const markdown = `# Main Title

## Introduction

This links to [Section 1](#section-1) and [Section 2](#section-2).

## Section 1

Content for section 1.

### Subsection 1.1

More content.

## Section 2

Content for section 2.
`;

      const result = await converter.markdownToDocx(markdown, {
        tocGeneration: true,
        preserveLinks: true,
      });

      expect(result.success).toBe(true);
      expect(result.metadata!.internalLinkCount).toBeGreaterThan(0);
    });

    test('should handle different templates', async () => {
      const markdown = '# Test\n\nSimple content.';
      const templates = converter.getAvailableTemplates();

      for (const template of templates.slice(0, 3)) { // Test first 3 templates
        const result = await converter.markdownToDocx(markdown, { template });
        expect(result.success).toBe(true);
      }
    });
  });

  describe('File Operations', () => {
    test('should convert markdown file to DOCX file', async () => {
      const markdown = '# File Test\n\nThis is a test from file.';
      const inputFile = path.join(testDataDir, 'test-input.md');
      const outputFile = path.join(outputDir, 'test-output.docx');

      await fs.writeFile(inputFile, markdown);

      const result = await converter.markdownFileToDocx(inputFile, outputFile);

      expect(result.success).toBe(true);
      expect(await fs.pathExists(outputFile)).toBe(true);
    });

    test('should handle invalid input file', async () => {
      const result = await converter.markdownFileToDocx(
        'non-existent-file.md',
        'output.docx'
      );

      expect(result.success).toBe(false);
      expect(result.error?.code).toBe('FILE_CONVERSION_FAILED');
    });
  });

  describe('Validation and Statistics', () => {
    test('should validate markdown content', async () => {
      const validMarkdown = `# Valid Document

## Section 1

Content with [link](https://example.com).

\`\`\`mermaid
graph LR
    A --> B
\`\`\`
`;

      const result = await converter.validateMarkdown(validMarkdown);

      expect(result.isValid).toBe(true);
      expect(result.suggestions.length).toBeGreaterThan(0); // Should suggest about Mermaid
    });

    test('should generate file statistics', async () => {
      const markdown = `# Statistics Test

This is a test document with multiple paragraphs.

![Image](image.png)

[Link](https://example.com)
`;
      const testFile = path.join(testDataDir, 'stats-test.md');
      await fs.writeFile(testFile, markdown);

      const stats = await converter.getConversionStats(testFile);

      expect(stats.wordCount).toBeGreaterThan(0);
      expect(stats.lineCount).toBeGreaterThan(0);
      expect(stats.imageCount).toBe(1);
      expect(stats.linkCount).toBe(1);
      expect(stats.readingTime).toBeGreaterThan(0);
    });
  });

  describe('Error Handling', () => {
    test('should handle empty markdown content', async () => {
      const result = await converter.markdownToDocx('');

      expect(result.success).toBe(false);
      expect(result.error?.message).toContain('Empty markdown content');
    });

    test('should handle malformed Mermaid diagrams', async () => {
      const markdown = `# Malformed Diagram

\`\`\`mermaid
invalid mermaid syntax here
\`\`\`
`;

      const result = await converter.markdownToDocx(markdown);

      // Should still succeed but may have warnings
      expect(result.success).toBe(true);
    });
  });

  describe('Performance', () => {
    test('should convert large documents efficiently', async () => {
      // Generate large markdown content
      const sections = Array.from({ length: 50 }, (_, i) => `
## Section ${i + 1}

Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.

\`\`\`javascript
function section${i + 1}() {
  return "Generated content for section ${i + 1}";
}
\`\`\`

| Column A | Column B | Column C |
|----------|----------|----------|
| Value ${i * 3 + 1} | Value ${i * 3 + 2} | Value ${i * 3 + 3} |
`).join('\n');

      const largeMarkdown = `# Large Document Test\n\n${sections}`;

      const startTime = Date.now();
      const result = await converter.markdownToDocx(largeMarkdown);
      const elapsed = Date.now() - startTime;

      expect(result.success).toBe(true);
      expect(elapsed).toBeLessThan(10000); // Should complete within 10 seconds
      expect(result.metadata!.processingTime).toBeLessThan(10000);
    }, 15000);
  });
});
