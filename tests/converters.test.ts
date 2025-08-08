// Mocks for dependencies
jest.mock('../src/processors/mermaid-png-processor', () => ({
  MermaidPNGProcessor: jest.fn().mockImplementation(() => ({
    processContent: jest.fn(async (content: string) => ({ content, diagramCount: 0, images: [] })),
    cleanup: jest.fn(),
  })),
}));

jest.mock('puppeteer', () => ({
  launch: jest.fn(),
}));

import { MarkdownToDocxConverter } from '../src/converters/markdown-to-docx';
import { DocxToMarkdownConverter } from '../src/converters/docx-to-markdown';

describe('Converters', () => {
  test('MarkdownToDocxConverter.stripFrontMatter parses metadata', () => {
    const converter = new MarkdownToDocxConverter() as any;
    const md = `---\ntitle: Test\nauthor: Alice\n---\n\n# Heading`;
    const { content, meta } = converter.stripFrontMatter(md);
    expect(meta).toEqual({ title: 'Test', author: 'Alice' });
    expect(content.trim()).toBe('# Heading');
  });

  test('DocxToMarkdownConverter.githubSlug creates slugs', () => {
    const converter = new DocxToMarkdownConverter() as any;
    expect(converter.githubSlug('Hello World!')).toBe('hello-world');
    expect(converter.githubSlug('Café au lait')).toBe('cafe-au-lait');
  });
});
