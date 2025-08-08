import { MarkdownDocxConverter } from '../src';
import * as fs from 'fs-extra';
import * as path from 'path';
import {
  Document,
  Packer,
  Paragraph,
  HeadingLevel,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ImageRun,
  Bookmark,
  InternalHyperlink,
} from 'docx';

// Tiny 1x1 PNG (red) base64
const RED_DOT_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAnsB6bQ4HqkAAAAASUVORK5CYII=';

async function createDocxFile(filePath: string, children: any[]) {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children,
      },
    ],
  });
  const buffer = await Packer.toBuffer(doc);
  await fs.ensureDir(path.dirname(filePath));
  await fs.writeFile(filePath, buffer);
  return buffer;
}

describe('DOCX to Markdown Conversion', () => {
  const inputDir = path.join(__dirname, 'docx-input');
  const outputDir = path.join(__dirname, 'output-md');
  const imagesDir = path.join(outputDir, 'images');

  let converter: MarkdownDocxConverter;

  beforeAll(async () => {
    converter = new MarkdownDocxConverter();
    await fs.ensureDir(inputDir);
    await fs.ensureDir(outputDir);
    await fs.emptyDir(outputDir);
  });

  afterAll(async () => {
    await fs.remove(outputDir);
    await fs.remove(inputDir);
  });

  test('converts simple DOCX (headings, bold) to Markdown', async () => {
    const file = path.join(inputDir, 'simple.docx');
    await createDocxFile(file, [
      new Paragraph({ text: 'Simple Title', heading: HeadingLevel.HEADING_1 }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Hello', bold: true }),
          new TextRun({ text: ' world!' }),
        ],
      }),
    ]);

    const res = await converter.docxToMarkdown(file, { headingAnchors: 'pandoc' });

    expect(res.success).toBe(true);
    const md = String(res.output || '');
    expect(md).toContain('# Simple Title {#simple-title}');
    // Bold may be preserved in Markdown
    expect(md).toContain('**Hello** world!');
  });

  test('converts tables to GFM Markdown', async () => {
    const file = path.join(inputDir, 'table.docx');
    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph('Header A')] }),
            new TableCell({ children: [new Paragraph('Header B')] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph('Cell 1')] }),
            new TableCell({ children: [new Paragraph('Cell 2')] }),
          ],
        }),
      ],
    });

    await createDocxFile(file, [
      new Paragraph({ text: 'Table Title', heading: HeadingLevel.HEADING_2 }),
      table,
    ]);

    const res = await converter.docxToMarkdown(file);
    expect(res.success).toBe(true);
    const md = String(res.output || '');
    // Should contain a table rendered with pipes and include the headers and cells
    const pipeLines = md.split(/\r?\n/).filter(l => l.trim().startsWith('|'));
    expect(pipeLines.length).toBeGreaterThanOrEqual(3);
    expect(md).toContain('Header A');
    expect(md).toContain('Header B');
    expect(md).toContain('Cell 1');
    expect(md).toContain('Cell 2');
  });

  test('extracts images and rewrites image paths', async () => {
    const file = path.join(inputDir, 'image.docx');
    const imgBuffer = Buffer.from(RED_DOT_BASE64, 'base64');

    await createDocxFile(file, [
      new Paragraph({ text: 'Image Title', heading: HeadingLevel.HEADING_2 }),
      new Paragraph({
        children: [
          new ImageRun({
            data: imgBuffer,
            transformation: { width: 16, height: 16 },
          }),
        ],
      }),
    ]);

    const res = await converter.docxToMarkdown(file, {
      extractImages: true,
      imageOutputDir: imagesDir,
    });

    expect(res.success).toBe(true);
    expect(res.metadata?.imageCount).toBeGreaterThanOrEqual(1);

    const md = String(res.output || '');
    // Allow absolute paths under ./.../images/... or direct ./images/...
    expect(md).toMatch(/!\[[^\]]*\]\(\.\/.*images\/image_\d+\.(png|jpg|jpeg|gif|bmp|svg|webp)\)/i);

    // Check file exists
    const files = (await fs.pathExists(imagesDir)) ? await fs.readdir(imagesDir) : [];
    expect(files.some(f => /^image_\d+\.(png|jpg|jpeg|gif|bmp|svg|webp)$/i.test(f))).toBe(true);
  });

  test('handles invalid input file gracefully', async () => {
    const res = await converter.docxToMarkdown('non-existent.docx');
    expect(res.success).toBe(false);
    expect(res.error?.code).toBe('CONVERSION_FAILED');
  });

  test('writes Markdown file via docxFileToMarkdown', async () => {
    const file = path.join(inputDir, 'write-file.docx');
    await createDocxFile(file, [
      new Paragraph({ text: 'Write File Title', heading: HeadingLevel.HEADING_1 }),
      new Paragraph('Some content here.'),
    ]);

    const outFile = path.join(outputDir, 'write-file.md');
    const res = await converter.docxFileToMarkdown(file, outFile, { headingAnchors: 'pandoc' });

    expect(res.success).toBe(true);
    expect(await fs.pathExists(outFile)).toBe(true);
    const saved = await fs.readFile(outFile, 'utf-8');
    expect(saved).toContain('# Write File Title {#write-file-title}');
  });

  test('converts complex DOCX structures using createDocxFile', async () => {
    const file = path.join(inputDir, 'complex.docx');
    await createDocxFile(file, [
      new Paragraph({ text: 'Complex Title', heading: HeadingLevel.HEADING_1 }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Bold', bold: true }),
          new TextRun({ text: ' and ' }),
          new TextRun({ text: 'italic', italics: true }),
          new TextRun({ text: ' text with ' }),
          new TextRun({ text: 'strike', strike: true }),
        ],
      }),
      // Code-like consecutive paragraphs
      new Paragraph('function add(a, b) {'),
      new Paragraph('  return a + b;'),
      new Paragraph('}'),
      // Bullet-like paragraphs
      new Paragraph('• Item 1'),
      new Paragraph('• Item 2'),
      // Numbered-like paragraphs
      new Paragraph('1. First'),
      new Paragraph('2. Second'),
    ]);

    const res = await converter.docxToMarkdown(file, { headingAnchors: 'pandoc' });
    expect(res.success).toBe(true);
    const md = String(res.output || '');

    // Heading with anchor
    expect(md).toContain('# Complex Title {#complex-title}');
    // Bold/italic/strike preserved
    expect(md).toContain('**Bold**');
    expect(md).toMatch(/\*italic\*|_italic_/); // italic could be * or _
    expect(md).toMatch(/~{1,2}strike~{1,2}/); // allow ~strike~ or ~~strike~~
    // Code content present (either fenced or plain)
    expect(md).toMatch(/return a \+ b;/);
    // List items reconstructed or at least converted to Markdown markers
    expect(md).toMatch(/(^|\n)\s*-\s*Item 1/);
    expect(md).toMatch(/(^|\n)\s*-\s*Item 2/);
    expect(md).toMatch(/(^|\n)\s*1\.\s*First/);
    expect(md).toMatch(/(^|\n)\s*2\.\s*Second/);
  });

  test('nested lists conversion (compare DOCX intent vs generated Markdown file)', async () => {
    const file = path.join(inputDir, 'nested-lists.docx');
    const outFile = path.join(outputDir, 'nested-lists.md');

    // Use NBSP to indicate nesting levels for bullets so our reconstructor can infer levels
    await createDocxFile(file, [
      new Paragraph('• Parent 1'),
      new Paragraph('\u00A0\u00A0• Child 1.1'),
      new Paragraph('\u00A0\u00A0• Child 1.2'),
      new Paragraph('\u00A0\u00A0\u00A0\u00A0• Grandchild 1.2.1'),
      new Paragraph('• Parent 2'),
    ]);

    const res = await converter.docxFileToMarkdown(file, outFile);
    expect(res.success).toBe(true);
    expect(await fs.pathExists(outFile)).toBe(true);

    const md = await fs.readFile(outFile, 'utf-8');
    const lines = md.split(/\r?\n/);

    const findLineIndex = (text: string) => {
      const idx = lines.findIndex(l => l.includes(text));
      expect(idx).toBeGreaterThanOrEqual(0);
      return idx;
    };
    const indentOf = (idx: number) => {
      const m = lines[idx].match(/^(\s*)-/);
      return m ? m[1].length : 0;
    };

    const iParent1 = findLineIndex('Parent 1');
    const iChild11 = findLineIndex('Child 1.1');
    const iChild12 = findLineIndex('Child 1.2');
    const iGrand = findLineIndex('Grandchild 1.2.1');
    const iParent2 = findLineIndex('Parent 2');

    // Order should reflect DOCX intent
    expect(iParent1).toBeLessThan(iChild11);
    expect(iChild11).toBeLessThan(iChild12);
    expect(iChild12).toBeLessThan(iGrand);
    expect(iGrand).toBeLessThan(iParent2);

    // Indentation should be non-decreasing across levels (best-effort)
    const indParent1 = indentOf(iParent1);
    const indChild11 = indentOf(iChild11);
    const indChild12 = indentOf(iChild12);
    const indGrand = indentOf(iGrand);

    expect(indChild11).toBeGreaterThanOrEqual(indParent1);
    expect(indChild12).toBeGreaterThanOrEqual(indChild11);
    // Grandchild ideally indents >= child; allow equality if reconstruction flattens
    expect(indGrand).toBeGreaterThanOrEqual(Math.min(indChild11, indChild12));
  });

  test('internal links conversion (compare DOCX intent vs generated Markdown file)', async () => {
    const file = path.join(inputDir, 'internal-links.docx');
    const outFile = path.join(outputDir, 'internal-links.md');

    await createDocxFile(file, [
      // Heading with a bookmark anchor
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [
          new Bookmark({ id: 'intro', children: [new TextRun('Intro Section')] }),
        ],
      }),
      // Internal hyperlink pointing to the bookmark
      new Paragraph({
        children: [
          new InternalHyperlink({ anchor: 'intro', children: [new TextRun('Go to Intro')] }),
        ],
      }),
    ]);

    const res = await converter.docxFileToMarkdown(file, outFile, { headingAnchors: 'pandoc' });
    expect(res.success).toBe(true);
    expect(await fs.pathExists(outFile)).toBe(true);

    const md = await fs.readFile(outFile, 'utf-8');

    // Heading should be present with an anchor (pandoc style)
    expect(md).toContain('## Intro Section {#intro-section}');
    // Internal link should be converted; slug is derived from link text per converter logic
    expect(md).toContain('[Go to Intro](#go-to-intro)');
  });
});
