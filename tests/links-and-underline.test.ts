import { MarkdownDocxConverter } from '../src';
import * as mammoth from 'mammoth';

describe('Markdown → DOCX: underline and nested link handling', () => {
  let converter: MarkdownDocxConverter;

  beforeAll(() => {
    converter = new MarkdownDocxConverter();
  });

  test('converts HTML underline <u>…</u> without emitting literal tags in DOCX→HTML', async () => {
    const md = `# Underline Test\n\n1. <u>Structurer le suivi des projets</u> de test.`;

    const result = await converter.markdownToDocx(md, { template: 'professional-report' });
    expect(result.success).toBe(true);
    const buffer = result.output as Buffer;

    const { value: html } = await mammoth.convertToHtml(buffer as any);
    expect(html).toMatch(/Structurer le suivi des projets/i);
  // Mammoth may drop underline styling; but literal <u> tags must not be present in output
  expect(html).not.toMatch(/<u>|<\/u>|&lt;u&gt;|&lt;\/u&gt;/i);
  }, 15000);

  test('normalizes malformed nested links [A]([B](URL)) to a single external link with text A', async () => {
    const md = `Intro [ITIL V4 Foundation]([ITIL Foundation: ITIL 4 Edition](https://example.com/itil.pdf)) end.`;

    const result = await converter.markdownToDocx(md, { template: 'professional-report', preserveLinks: true });
    expect(result.success).toBe(true);
    const buffer = result.output as Buffer;

    const { value: html } = await mammoth.convertToHtml(buffer as any);
    // Expect a single anchor with the normalized href and the outer text as label
    expect(html).toMatch(/<a[^>]+href=\"https:\/\/example\.com\/itil\.pdf\"[^>]*>\s*ITIL V4 Foundation\s*<\/a>/i);
  }, 15000);
});
