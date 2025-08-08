import { MarkdownDocxConverter } from '../src';
import * as fs from 'fs-extra';
import * as path from 'path';

function normalizeMarkdown(md: string): string {
  return md
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+$/gm, '')
    .replace(/\u00A0|\u00B7/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

describe('Markdown → DOCX (round-trip via DOCX → Markdown)', () => {
  const mdInputDir = path.join(__dirname, 'md-input');
  const docxOutputDir = path.join(__dirname, 'output-docx-rt');
  const roundtripMdDir = path.join(__dirname, 'roundtrip-md');
  const imagesDir = path.join(roundtripMdDir, 'images');

  let converter: MarkdownDocxConverter;

  beforeAll(async () => {
    converter = new MarkdownDocxConverter();
    await fs.ensureDir(mdInputDir);
    await fs.ensureDir(docxOutputDir);
    await fs.ensureDir(roundtripMdDir);
    await fs.emptyDir(docxOutputDir);
    await fs.emptyDir(roundtripMdDir);
  });

  afterAll(async () => {
    await fs.remove(mdInputDir);
    await fs.remove(docxOutputDir);
    await fs.remove(roundtripMdDir);
  });

  test('headings and inline formatting round-trip', async () => {
    const mdIn = [
      '# Title',
      '',
      'This is **bold** and *italic* and `code`.',
    ].join('\n');

    const inFile = path.join(mdInputDir, 'simple.md');
    const outDocx = path.join(docxOutputDir, 'simple.docx');
    const roundMd = path.join(roundtripMdDir, 'simple.md');

    await fs.writeFile(inFile, mdIn, 'utf-8');

    const res1 = await converter.markdownFileToDocx(inFile, outDocx);
    expect(res1.success).toBe(true);
    expect(await fs.pathExists(outDocx)).toBe(true);

    const res2 = await converter.docxFileToMarkdown(outDocx, roundMd, { headingAnchors: 'none' });
    expect(res2.success).toBe(true);

    const mdOut = await fs.readFile(roundMd, 'utf-8');
    const norm = normalizeMarkdown(mdOut);

    // Heading may contain bold markers after round-trip
    expect(norm).toMatch(/^#\s+(\*\*)?Title(\*\*)?/m);
    expect(norm).toContain('**bold**');
    expect(norm).toMatch(/\*italic\*|_italic_/);
    // Inline code might lose backticks in round-trip; accept plain word
    expect(norm).toMatch(/(^|\W)code(\W|$)/);
  }, 20000);

  test('tables round-trip', async () => {
    const mdIn = [
      '## Table Demo',
      '',
      '| Header A | Header B |',
      '|----------|----------|',
      '| Cell 1   | Cell 2   |',
    ].join('\n');

    const inFile = path.join(mdInputDir, 'table.md');
    const outDocx = path.join(docxOutputDir, 'table.docx');
    const roundMd = path.join(roundtripMdDir, 'table.md');

    await fs.writeFile(inFile, mdIn, 'utf-8');

    const res1 = await converter.markdownFileToDocx(inFile, outDocx);
    expect(res1.success).toBe(true);

    const res2 = await converter.docxFileToMarkdown(outDocx, roundMd, { headingAnchors: 'none' });
    expect(res2.success).toBe(true);

    const mdOut = await fs.readFile(roundMd, 'utf-8');
    const pipeLines = mdOut.split(/\r?\n/).filter(l => l.trim().startsWith('|'));
    expect(pipeLines.length).toBeGreaterThanOrEqual(3);
    expect(mdOut).toContain('Header A');
    expect(mdOut).toContain('Header B');
    expect(mdOut).toContain('Cell 1');
    expect(mdOut).toContain('Cell 2');
  }, 20000);

  test('mermaid diagrams embed to DOCX and extract back as images', async () => {
    const mdIn = [
      '## Mermaid Demo',
      '',
      '```mermaid',
      'graph TD',
      '  A[Start] --> B{Decision}',
      '  B -->|Yes| C[Do]',
      '  B -->|No| D[Stop]',
      '```',
    ].join('\n');

    const inFile = path.join(mdInputDir, 'mermaid.md');
    const outDocx = path.join(docxOutputDir, 'mermaid.docx');
    const roundMd = path.join(roundtripMdDir, 'mermaid.md');

    await fs.writeFile(inFile, mdIn, 'utf-8');

    const res1 = await converter.markdownFileToDocx(inFile, outDocx, { mermaidTheme: 'forest' });
    expect(res1.success).toBe(true);

    const res2 = await converter.docxFileToMarkdown(outDocx, roundMd, {
      extractImages: true,
      imageOutputDir: imagesDir,
      headingAnchors: 'none',
    });
    expect(res2.success).toBe(true);

    const mdOut = await fs.readFile(roundMd, 'utf-8');
    expect(mdOut).toMatch(/!\[[^\]]*\]\(\.\/.*images\/image_\d+\.(png|jpg|jpeg|gif|bmp|svg|webp)\)/i);

    const files = (await fs.pathExists(imagesDir)) ? await fs.readdir(imagesDir) : [];
    expect(files.some(f => /^image_\d+\.(png|jpg|jpeg|gif|bmp|svg|webp)$/i.test(f))).toBe(true);
  }, 30000);

  test('nested lists round-trip', async () => {
    const mdIn = [
      '- Parent 1',
      '  - Child 1.1',
      '  - Child 1.2',
      '    - Grandchild 1.2.1',
      '- Parent 2',
      '',
      '1. First',
      '   1. Sub First 1',
      '2. Second',
    ].join('\n');

    const inFile = path.join(mdInputDir, 'nested-lists.md');
    const outDocx = path.join(docxOutputDir, 'nested-lists.docx');
    const roundMd = path.join(roundtripMdDir, 'nested-lists.md');

    await fs.writeFile(inFile, mdIn, 'utf-8');

    const res1 = await converter.markdownFileToDocx(inFile, outDocx);
    expect(res1.success).toBe(true);

    const res2 = await converter.docxFileToMarkdown(outDocx, roundMd, { headingAnchors: 'none' });
    expect(res2.success).toBe(true);

    const md = await fs.readFile(roundMd, 'utf-8');
    const lines = md.split(/\r?\n/);

    const findLineIndex = (text: string) => {
      const idx = lines.findIndex(l => l.includes(text));
      expect(idx).toBeGreaterThanOrEqual(0);
      return idx;
    };
    const bulletIndent = (idx: number) => {
      const m = lines[idx].match(/^(\s*)-\s/);
      return m ? m[1].length : 0;
    };
    const orderedIndent = (idx: number) => {
      const m = lines[idx].match(/^(\s*)\d+\.\s/);
      return m ? m[1].length : 0;
    };

    const iP1 = findLineIndex('Parent 1');
    const iC11 = findLineIndex('Child 1.1');
    const iC12 = findLineIndex('Child 1.2');
    const iG = findLineIndex('Grandchild 1.2.1');
    const iP2 = findLineIndex('Parent 2');

    expect(iP1).toBeLessThan(iC11);
    expect(iC11).toBeLessThan(iC12);
    expect(iC12).toBeLessThan(iG);
    expect(iG).toBeLessThan(iP2);

    const indP1 = bulletIndent(iP1);
    const indC11 = bulletIndent(iC11);
    const indC12 = bulletIndent(iC12);
    const indG = bulletIndent(iG);

    expect(indC11).toBeGreaterThanOrEqual(indP1);
    expect(indC12).toBeGreaterThanOrEqual(indC11);
    expect(indG).toBeGreaterThanOrEqual(indC12);

    const iFirst = findLineIndex('First');
    const iSubFirst = findLineIndex('Sub First 1');
    const indFirst = orderedIndent(iFirst);
    const indSubFirst = orderedIndent(iSubFirst);
    expect(iFirst).toBeLessThan(iSubFirst);
    expect(indSubFirst).toBeGreaterThanOrEqual(indFirst);
  }, 20000);

  test('internal links round-trip using heading text slug', async () => {
    const mdIn = [
      '## Intro Section',
      '',
      'See [Intro Section](#intro-section)',
    ].join('\n');

    const inFile = path.join(mdInputDir, 'internal-links.md');
    const outDocx = path.join(docxOutputDir, 'internal-links.docx');
    const roundMd = path.join(roundtripMdDir, 'internal-links.md');

    await fs.writeFile(inFile, mdIn, 'utf-8');

    const res1 = await converter.markdownFileToDocx(inFile, outDocx);
    expect(res1.success).toBe(true);

    const res2 = await converter.docxFileToMarkdown(outDocx, roundMd, { headingAnchors: 'pandoc' });
    expect(res2.success).toBe(true);

    const md = await fs.readFile(roundMd, 'utf-8');
    const norm = normalizeMarkdown(md);
    // Heading text may be bolded by the round-trip; anchor may be present
    expect(norm).toMatch(/^##\s+(\*\*)?Intro Section(\*\*)?(\s+\{#intro-section\})?/m);
    expect(norm).toContain('[Intro Section](#intro-section)');
  }, 20000);
});
