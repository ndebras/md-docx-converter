import { MermaidPNGProcessor } from '../src/processors/mermaid-png-processor';

describe('MermaidPNGProcessor Offline Rendering', () => {
  test('renders diagram without network access', async () => {
    const processor = new MermaidPNGProcessor();
    const markdown = '```mermaid\ngraph LR\nA-->B\n```';

    const result = await processor.processContent(markdown);

    expect(result.diagramCount).toBe(1);
    expect(result.images.length).toBe(1);
    expect(result.images[0].buffer.length).toBeGreaterThan(0);

    await processor.cleanup();
    processor.cleanupTempDir();
  });
});
