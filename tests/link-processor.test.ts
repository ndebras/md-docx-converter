import { LinkProcessor } from '../src/processors/link-processor';
import { ProcessedLink } from '../src/types';

describe('LinkProcessor', () => {
  test('processMarkdownLinks extracts links and headings', () => {
    const processor = new LinkProcessor();
    const md = '# Intro\nSee [Intro](#intro) and [Google](http://google.com)';
    const { links } = processor.processMarkdownLinks(md);

    expect(links).toHaveLength(2);
    const internal = links.find(l => l.type === 'anchor');
    const external = links.find(l => l.type === 'external');

    expect(internal).toMatchObject({ text: 'Intro', url: '#intro', isValid: true });
    expect(external).toMatchObject({ text: 'Google', url: 'http://google.com', isValid: true });
  });

  test('convertToWordLinks maps link types correctly', () => {
    const processor = new LinkProcessor();
    const links: ProcessedLink[] = [
      { text: 'Google', url: 'http://google.com', type: 'external', isValid: true },
      { text: 'Intro', url: '#intro', type: 'anchor', isValid: true },
    ];

    const wordLinks = processor.convertToWordLinks(links);
    expect(wordLinks).toEqual([
      { text: 'Google', url: 'http://google.com', type: 'hyperlink', style: 'Hyperlink' },
      { text: 'Intro', url: 'intro', type: 'bookmark', style: 'IntenseReference' },
    ]);
  });

  test('validateLinks identifies invalid URLs', () => {
    const processor = new LinkProcessor();
    const links: ProcessedLink[] = [
      { text: 'Good', url: 'http://example.com', type: 'external', isValid: true },
      { text: 'Mail', url: 'mailto:test@example.com', type: 'external', isValid: true },
      { text: 'Bad', url: 'javascript:alert(1)', type: 'external', isValid: true },
    ];

    const validated = processor.validateLinks(links);
    expect(validated[0].isValid).toBe(true);
    expect(validated[1].isValid).toBe(false);
    expect(validated[2].isValid).toBe(false);
  });
});
