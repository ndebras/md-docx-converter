import * as mammoth from 'mammoth';
import TurndownService from 'turndown';
// GFM plugin for tables, strikethrough, task lists
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import { gfm } from 'turndown-plugin-gfm';
import * as fs from 'fs-extra';
import * as path from 'path';
import { DocxToMarkdownOptions, ConversionResult } from '../types';
import { FileUtils, StringUtils, PerformanceUtils } from '../utils';
import { logger } from '../utils/logger';

/**
 * Advanced DOCX to Markdown converter with formatting preservation
 */
export class DocxToMarkdownConverter {
  private turndownService: TurndownService;
  private extractedImages: Array<{ filename: string; buffer: Buffer }> = [];

  constructor(options: DocxToMarkdownOptions = {}) {
    this.turndownService = this.configureTurndown(options);
  }

  /**
   * Convert DOCX to Markdown
   */
  async convert(
    docxFilePath: string,
    options: DocxToMarkdownOptions = {}
  ): Promise<ConversionResult> {
    const startTime = Date.now();

    try {
      logger.info('Starting DOCX to Markdown conversion', { file: docxFilePath });

      // Validate input file
      const validation = FileUtils.validateFile(docxFilePath, ['.docx']);
      if (!validation.isValid) {
        throw new Error(`Invalid input file: ${validation.errors.map(e => e.message).join(', ')}`);
      }

      // Read DOCX file
      const { result: docxBuffer, elapsed: readTime } = 
        await PerformanceUtils.measureAsync(async () => {
          return await fs.readFile(docxFilePath);
        }, 'file-read');

      logger.debug(`File read completed in ${readTime}ms`);

      // Configure mammoth options
      const mammothOptions = this.configureMammoth(options);

      // Extract content and images
      const { result: extraction, elapsed: extractTime } = 
        await PerformanceUtils.measureAsync(async () => {
          return await this.extractDocxContent(docxBuffer, mammothOptions, options);
        }, 'content-extraction');

      logger.debug(`Content extraction completed in ${extractTime}ms`);

      // Convert HTML to Markdown unless we already have Markdown (fallback)
      let markdown: string;
      let conversionTime = 0;
      if (extraction.isMarkdown) {
        markdown = extraction.html;
        logger.warn('Using fallback Markdown extracted from DOCX');
      } else {
        try {
          const conversion = await PerformanceUtils.measureAsync(async () => {
            return this.convertHtmlToMarkdown(extraction.html, options);
          }, 'html-to-markdown');
          markdown = conversion.result;
          conversionTime = conversion.elapsed;
          logger.debug(`HTML to Markdown conversion completed in ${conversionTime}ms`);
        } catch (turndownError) {
          logger.error('HTML to Markdown conversion failed, falling back to raw text', { error: turndownError });
          const raw = await mammoth.extractRawText({ buffer: docxBuffer as any });
          markdown = raw.value;
          // mark that this is degraded output
          extraction.messages.push({ type: 'warning', message: 'Fell back to raw text due to HTML->Markdown conversion error' });
        }
      }

      // Save extracted images if requested (skip when using raw-text fallback where no images are collected)
      if (options.extractImages && options.imageOutputDir && this.extractedImages.length > 0 && !extraction.isMarkdown) {
        await this.saveExtractedImages(options.imageOutputDir);
      } else if (extraction.isMarkdown) {
        // ensure no stale images are counted
        this.extractedImages = [];
      }

      // Post-process markdown
  const processedMarkdown = this.postProcessMarkdown(markdown || '', options);

      const totalTime = Date.now() - startTime;

  const metadata = {
        inputSize: docxBuffer.length,
        outputSize: Buffer.from(processedMarkdown, 'utf-8').length,
        processingTime: totalTime,
        imageCount: this.extractedImages.length,
      };

      logger.info('DOCX to Markdown conversion completed successfully', metadata);

  const warnings = this.getConversionWarnings(extraction.messages, options);
  return {
        success: true,
        output: processedMarkdown,
        metadata,
    warnings,
      };

    } catch (error) {
      logger.error('DOCX to Markdown conversion failed', { error, file: docxFilePath });

      return {
        success: false,
        error: {
          code: 'CONVERSION_FAILED',
          message: error instanceof Error ? error.message : String(error),
          details: { 
            file: docxFilePath,
            processingTime: Date.now() - startTime,
          },
        },
      };
    } finally {
      this.cleanup();
    }
  }

  /**
   * Configure Turndown service
   */
  private configureTurndown(options: DocxToMarkdownOptions): TurndownService {
    const turndown = new TurndownService({
      headingStyle: 'atx',
      hr: '---',
      bulletListMarker: '-',
      codeBlockStyle: 'fenced',
      fence: '```',
      emDelimiter: '*',
      strongDelimiter: '**',
      linkStyle: 'inlined',
      linkReferenceStyle: 'full',
    });

    // Keep HTML for underline/superscript/subscript when present in source
    try {
      // keep preserves tags as inline HTML
      (turndown as any).keep?.(['u', 'sub', 'sup']);
    } catch (_) {
      // no-op if keep not available
    }

    // Custom rules for better formatting
    this.addCustomTurndownRules(turndown, options);

    // Enable GitHub Flavored Markdown features (tables, strikethrough, task lists)
    try {
      turndown.use(gfm);
    } catch (e) {
      logger.warn('Failed to enable GFM plugin for Turndown', { error: e });
    }

    return turndown;
  }

  /**
   * Add custom Turndown rules
   */
  private addCustomTurndownRules(turndown: TurndownService, options: DocxToMarkdownOptions): void {
    // Preserve code blocks: handle <pre> and <pre><code>
    turndown.addRule('codeBlock', {
      filter: (node: any) => {
        const name = (node.nodeName || node.tagName || '').toUpperCase();
        return name === 'PRE';
      },
      replacement: (_content: string, node: any) => {
        const getText = (n: any) => (n && typeof n.textContent === 'string') ? n.textContent : '';
        const codeChild = node && node.querySelector ? node.querySelector('code') : null;
        const text = codeChild ? getText(codeChild) : getText(node);
        const content = (text || '').replace(/\r\n?/g, '\n');
        let language = '';
        if (content.includes('function') && content.includes('{')) language = 'javascript';
        else if (content.includes('def ') && content.includes(':')) language = 'python';
        else if (/^\s*#include/m.test(content)) language = 'c';
        return `\n\`\`\`${language}\n${content}\n\`\`\`\n`;
      },
    });

    // Preserve inline code
    turndown.addRule('inlineCode', {
      filter: ['code'],
      replacement: (content) => {
        return content.includes('`') ? `\`\`${content}\`\`` : `\`${content}\``;
      },
    });

    // Strikethrough
    turndown.addRule('strikethrough', {
      filter: ['s', 'del'],
      replacement: (content) => `~~${content}~~`,
    });

  // Rely on default handling for tables (avoid DOM-specific APIs)

    // Handle images
    turndown.addRule('image', {
      filter: 'img',
      replacement: (content, node: any) => {
        const getAttr = (n: any, name: string) => (typeof n.getAttribute === 'function' ? n.getAttribute(name) : (n.attributes?.[name] ?? '')) || '';
        const alt = getAttr(node, 'alt') || 'Image';
        const src = getAttr(node, 'src') || '';
        
        // Check if this is an extracted image
        const extractedImage = this.extractedImages.find(img => src.includes(img.filename));
        // Normalize image output path to POSIX-style relative path
        const normalizePath = (p: string) => {
          // Replace backslashes with slashes and trim leading './' or '.\\'
          let s = p.replace(/\\/g, '/').replace(/^\.\/?/, '');
          // Ensure a single leading './'
          return `./${s}`;
        };
        const outDir = options.imageOutputDir || 'images';
        const imagePath = extractedImage 
          ? normalizePath(`${outDir}/${extractedImage.filename}`)
          : normalizePath(src || '');

        return `![${alt}](${imagePath})`;
      },
    });

    // Handle hyperlinks
    turndown.addRule('hyperlink', {
      filter: (node: any) => {
        const name = node.nodeName || node.tagName || '';
        const getAttr = (n: any, attr: string) => (typeof n.getAttribute === 'function' ? n.getAttribute(attr) : (n.attributes?.[attr] ?? '')) || '';
        return name.toUpperCase() === 'A' && !!getAttr(node, 'href');
      },
      replacement: (content, node: any) => {
        const getAttr = (n: any, attr: string) => (typeof n.getAttribute === 'function' ? n.getAttribute(attr) : (n.attributes?.[attr] ?? '')) || '';
        const rawHref = getAttr(node, 'href') || '';
        // Internal anchors: normalize to heading-style slug so we don't rely on injected <a id> tags
        if (rawHref.startsWith('#')) {
          // Prefer slug derived from link text when available; fallback to sanitized href fragment
          const text = (content || '').trim();
          const normalized = this.normalizeInternalAnchor(rawHref, text);
          return `[${content}](${normalized})`;
        }
        // External links as-is
        return `[${content}](${rawHref})`;
      },
    });

    // Convert HTML tables to GFM Markdown tables (handles colspan/rowspan best-effort)
    turndown.addRule('tableToMarkdown', {
      filter: (node: any) => {
        const name = (node.nodeName || node.tagName || '').toUpperCase();
        return name === 'TABLE';
      },
      replacement: (_content: string, node: any) => {
        const qsa = (n: any, sel: string) => (n.querySelectorAll ? Array.from(n.querySelectorAll(sel)) : []);
        const rows: any[] = qsa(node, 'tr');
        if (!rows.length) return '';

        // Build grid honoring col/row spans
        const grid: string[][] = [];
        const cellTypes: ('th'|'td')[] = [];
        let maxCols = 0;

        const getAttr = (n: any, name: string) => (typeof n.getAttribute === 'function' ? n.getAttribute(name) : (n.attributes?.[name] ?? '')) || '';
        const cellText = (c: any) => (c.textContent || '').replace(/\r\n?/g, ' ').replace(/\|/g, '\\|').replace(/\s+/g, ' ').trim();

        rows.forEach((row: any, rIdx: number) => {
          if (!grid[rIdx]) grid[rIdx] = [];
          let cIdx = 0;
          const cells = qsa(row, 'th,td');
          cells.forEach((cell: any) => {
            // Find next free column in this row
            while (grid[rIdx][cIdx] !== undefined) cIdx++;
            const colspan = Math.max(1, parseInt(getAttr(cell, 'colspan') || '1', 10));
            const rowspan = Math.max(1, parseInt(getAttr(cell, 'rowspan') || '1', 10));
            const text = cellText(cell);
            const isTh = ((cell.nodeName || cell.tagName || '').toUpperCase() === 'TH');
            if (rIdx === 0) cellTypes[cIdx] = isTh ? 'th' : 'td';

            // Place text at [rIdx, cIdx]
            grid[rIdx][cIdx] = text;
            // Fill colspan span with empty placeholders to keep alignment
            for (let cs = 1; cs < colspan; cs++) grid[rIdx][cIdx + cs] = '';
            // Propagate into rowspans (duplicate text for readability)
            for (let rs = 1; rs < rowspan; rs++) {
              const rr = rIdx + rs;
              if (!grid[rr]) grid[rr] = [];
              grid[rr][cIdx] = text;
              for (let cs = 1; cs < colspan; cs++) grid[rr][cIdx + cs] = '';
            }
            cIdx += colspan;
            maxCols = Math.max(maxCols, cIdx);
          });
          maxCols = Math.max(maxCols, (grid[rIdx] || []).length);
        });

        if (!maxCols) return '';
        // Pad rows to maxCols
        grid.forEach((r, i) => {
          const row = r || (grid[i] = []);
          while (row.length < maxCols) row.push('');
        });

        // Heuristic: if table degenerated to 1 column but has many rows, try to reflow into multiple columns
        if (maxCols === 1 && grid.length >= 4) {
          const flat = grid.map(r => r[0]);
          const total = flat.length;
          const divisors: number[] = [];
          const maxTry = 12; // don't create too many columns
          for (let d = 2; d <= Math.min(maxTry, total); d++) {
            if (total % d === 0) divisors.push(d);
          }
          if (divisors.length > 0) {
            const target = Math.round(Math.sqrt(total));
            let best = divisors[0];
            let bestDelta = Math.abs(divisors[0] - target);
            for (const d of divisors) {
              const delta = Math.abs(d - target);
              if (delta < bestDelta) { best = d; bestDelta = delta; }
            }
            const cols = best;
            const rowsCount = total / cols;
            const reflow: string[][] = [];
            for (let r = 0; r < rowsCount; r++) {
              const row: string[] = [];
              for (let c = 0; c < cols; c++) {
                row.push(flat[r * cols + c] || '');
              }
              reflow.push(row);
            }
            grid.splice(0, grid.length, ...reflow);
            maxCols = cols;
          }
        }

        // Decide header row: use first row with any TH, else first row
  const hasAnyTh = cellTypes.some(t => t === 'th');
  const header = grid[0];
        const lines: string[] = [];
        const rowToLine = (r: string[]) => `| ${r.join(' | ')} |`;
        if (hasAnyTh) {
          lines.push(rowToLine(header));
        } else {
          // If header row is empty, synthesize blanks
          lines.push(rowToLine(header));
        }
        lines.push(`| ${Array(maxCols).fill('---').join(' | ')} |`);
        for (let i = 1; i < grid.length; i++) lines.push(rowToLine(grid[i]));
        return `\n${lines.join('\n')}\n`;
      },
    });

    // Handle blockquotes
    turndown.addRule('blockquote', {
      filter: 'blockquote',
      replacement: (content) => {
        return content
          .split('\n')
          .map(line => line.trim() ? `> ${line}` : '>')
          .join('\n') + '\n';
      },
    });

    // Handle horizontal rules
    turndown.addRule('horizontalRule', {
      filter: 'hr',
      replacement: () => '\n---\n',
    });

    // Preserve line breaks in certain contexts
    if (options.preserveFormatting) {
      turndown.addRule('lineBreak', {
        filter: 'br',
        replacement: () => '  \n', // Two spaces + newline for Markdown line break
      });
    }
  }

  /**
   * Configure mammoth options
   */
  private configureMammoth(options: DocxToMarkdownOptions): any {
    const styleMap = [
      // Headings
      "p[style-name='Heading 1'] => h1:fresh",
      "p[style-name='Heading 2'] => h2:fresh",
      "p[style-name='Heading 3'] => h3:fresh",
      "p[style-name='Heading 4'] => h4:fresh",
      "p[style-name='Heading 5'] => h5:fresh",
      "p[style-name='Heading 6'] => h6:fresh",
      
      // Built-in heading styles
      "p[style-name='heading 1'] => h1:fresh",
      "p[style-name='heading 2'] => h2:fresh",
      "p[style-name='heading 3'] => h3:fresh",
      "p[style-name='heading 4'] => h4:fresh",
      "p[style-name='heading 5'] => h5:fresh",
      "p[style-name='heading 6'] => h6:fresh",
      
      // Code blocks
      "p[style-name='Code'] => pre:separator('\\n')",
      "p[style-name='code'] => pre:separator('\\n')",
      
      // Blockquotes
      "p[style-name='Quote'] => blockquote > p:fresh",
      "p[style-name='quote'] => blockquote > p:fresh",

      // Normal paragraphs
      "p[style-name='Normal'] => p:fresh",
    ];

    const mammothOptions: any = {
      styleMap,
      includeDefaultStyleMap: true,
      includeEmbeddedStyleMap: true,
      ignoreEmptyParagraphs: false,
      preserveLineBreaks: options.preserveFormatting || false,
    };

    if (options.extractImages) {
      mammothOptions.convertImage = mammoth.images.imgElement(async (image: any) => {
        try {
          const extension = this.getImageExtension(image.contentType || 'image/png');
          const filename = `image_${this.extractedImages.length + 1}${extension}`;
          const buffer: Buffer = await image.read();
          this.extractedImages.push({ filename, buffer });
          logger.debug(`Extracted image: ${filename}`);
          return { src: filename, altText: image.altText || 'Extracted image' };
        } catch (err) {
          logger.warn('Failed to process image during conversion', { error: err });
          // Let Mammoth fall back to its default inline behavior
          return null as any;
        }
      });
    }

    return mammothOptions;
  }

  /**
   * Extract DOCX content
   */
  private async extractDocxContent(
    docxBuffer: Buffer,
    mammothOptions: any,
    options: DocxToMarkdownOptions
  ): Promise<{ html: string; messages: any[]; isMarkdown?: boolean }> {
    try {
      // In Node, pass a Buffer to Mammoth
      const result = await mammoth.convertToHtml({ buffer: docxBuffer } as any, mammothOptions);
      
      logger.debug('DOCX extraction completed', {
        messageCount: result.messages.length,
        warnings: result.messages.filter((m: any) => m.type === 'warning').length,
        errors: result.messages.filter((m: any) => m.type === 'error').length,
      });

      return {
        html: result.value,
        messages: result.messages,
      };
    } catch (error) {
      logger.error('Failed to extract DOCX content', {
        error: error instanceof Error ? { message: error.message, stack: error.stack } : error,
      });
      // Fallback to raw text when HTML path fails; this avoids downstream parser errors
      try {
        const raw = await mammoth.extractRawText({ buffer: docxBuffer } as any);
        logger.warn('Falling back to Mammoth.extractRawText');
        return {
          html: raw.value, // treat as already-markdown-ish plain text
          messages: [
            ...(raw.messages || []),
            { type: 'warning', message: 'Used raw text fallback; formatting may be lost; images will not be saved.' },
          ],
          isMarkdown: true,
        };
      } catch (fallbackError) {
        logger.error('Fallback to raw text conversion also failed', { 
          error: fallbackError instanceof Error ? { message: fallbackError.message, stack: fallbackError.stack } : fallbackError 
        });
        throw error;
      }
    }
  }

  /**
   * Handle image extraction
   */
  private async handleImageExtraction(image: any, options: DocxToMarkdownOptions): Promise<any> {
    if (!options.extractImages) {
      return image;
    }

    try {
      const extension = this.getImageExtension(image.contentType || 'image/png');
      const filename = `image_${this.extractedImages.length + 1}${extension}`;
      
      // Mammoth image API provides a read() function that returns a Promise for the buffer in Node
      let buffer: Buffer | undefined = undefined;
      if (typeof image.read === 'function') {
        buffer = await image.read();
      } else if (image.buffer) {
        buffer = image.buffer as Buffer;
      }
      if (!buffer) {
        logger.warn('Image buffer not available from Mammoth; skipping this image');
        return image;
      }
      
      this.extractedImages.push({ filename, buffer });

      logger.debug(`Extracted image: ${filename}`);
      
  return { src: filename, altText: image.altText || 'Extracted image' };
    } catch (error) {
      logger.warn('Failed to extract image', { error });
      return image;
    }
  }

  /**
   * Get image file extension from content type
   */
  private getImageExtension(contentType: string): string {
    const mimeToExt: Record<string, string> = {
      'image/png': '.png',
      'image/jpeg': '.jpg',
      'image/jpg': '.jpg',
      'image/gif': '.gif',
      'image/bmp': '.bmp',
      'image/svg+xml': '.svg',
      'image/webp': '.webp',
    };

    return mimeToExt[contentType] || '.png';
  }

  /**
   * Convert HTML to Markdown
   */
  private convertHtmlToMarkdown(html: string, options: DocxToMarkdownOptions): string {
    try {
      // Clean up HTML before conversion
      const cleanedHtml = this.cleanHtml(html, options);
      
      // Convert to Markdown
      const markdown = this.turndownService.turndown(cleanedHtml);
      
      return markdown;
    } catch (error) {
      logger.error('Failed to convert HTML to Markdown', { error });
      throw error;
    }
  }

  /**
   * Clean HTML for better Markdown conversion
   */
  private cleanHtml(html: string, options: DocxToMarkdownOptions): string {
    let cleaned = html;

    // Remove Word-specific namespaces and attributes
    cleaned = cleaned.replace(/\s*xmlns[^=]*="[^"]*"/g, '');
    cleaned = cleaned.replace(/\s*o:[^=]*="[^"]*"/g, '');
    cleaned = cleaned.replace(/\s*w:[^=]*="[^"]*"/g, '');
    cleaned = cleaned.replace(/\s*v:[^=]*="[^"]*"/g, '');

    // Clean up empty paragraphs and divs
    cleaned = cleaned.replace(/<p[^>]*>\s*<\/p>/g, '');
    cleaned = cleaned.replace(/<div[^>]*>\s*<\/div>/g, '');

  // Normalize whitespace between tags only (avoid damaging table/cell content)
  cleaned = cleaned.replace(/>\s+</g, '><');

    // Normalize inline styling to semantic tags for better Markdown conversion
    // underline -> <u>
    cleaned = cleaned.replace(/<span[^>]*style="[^"]*text-decoration\s*:\s*underline;?[^"]*"[^>]*>(.*?)<\/span>/gi, '<u>$1</u>');
    // superscript -> <sup>
    cleaned = cleaned.replace(/<span[^>]*style="[^"]*vertical-align\s*:\s*super;?[^"]*"[^>]*>(.*?)<\/span>/gi, '<sup>$1</sup>');
    // subscript -> <sub>
    cleaned = cleaned.replace(/<span[^>]*style="[^"]*vertical-align\s*:\s*sub;?[^"]*"[^>]*>(.*?)<\/span>/gi, '<sub>$1</sub>');
  // strikethrough -> <del>
  cleaned = cleaned.replace(/<span[^>]*style="[^"]*text-decoration\s*:\s*line-through;?[^"]*"[^>]*>(.*?)<\/span>/gi, '<del>$1</del>');


  // Do not attempt to rewrite lists; rely on Mammoth to produce proper <ul>/<ol>

  // Fix table formatting
  cleaned = this.fixTableFormatting(cleaned);

  // Reconstruct lists when Mammoth produced list-like paragraphs without proper <ul>/<ol>
  cleaned = this.reconstructListsIfMissing(cleaned);

    return cleaned;
  }

  /**
   * Fix list formatting in HTML
   */
  // Removed custom list formatting fixer; relying on Mammoth's list handling

  /**
   * Fix table formatting in HTML
   */
  private fixTableFormatting(html: string): string {
    // Ensure proper table structure
  let fixed = html;
    
    // Wrap orphaned td elements in tr
    fixed = fixed.replace(/(<td[^>]*>.*?<\/td>)(?!\s*<\/tr>)/g, '<tr>$1</tr>');
    
    // Wrap orphaned tr elements in tbody
    fixed = fixed.replace(/(<tr[^>]*>.*?<\/tr>)(?!\s*<\/tbody>)/g, '<tbody>$1</tbody>');
    
  // Merge adjacent tbodys into a single tbody
  fixed = fixed.replace(/<\/tbody>\s*<tbody>/g, '');

  // Remove paragraph wrappers inside table cells
  fixed = fixed.replace(/<td([^>]*)>\s*<p[^>]*>(.*?)<\/p>\s*<\/td>/g, '<td$1>$2</td>');
    
    return fixed;
  }

  private reconstructListsIfMissing(html: string): string {
    const paraRe = /<p([^>]*)>([\s\S]*?)<\/p>/gi;
    const items: Array<{ idx: number; raw: string; type: 'ul' | 'ol' | null; level: number; text: string; inLi: boolean }> = [];
    let match: RegExpExecArray | null;
    const paras: string[] = [];
    let lastIndex = 0;
    // Collect paragraphs
    while ((match = paraRe.exec(html)) !== null) {
      const before = html.slice(lastIndex, match.index);
      paras.push(before);
      const attrs = match[1] || '';
      const content = match[2] || '';
      const inLi = before.lastIndexOf('<li') > before.lastIndexOf('</li>');
      const parsed = this.parseListishParagraph(attrs, content, inLi);
      items.push({ idx: paras.length - 1, raw: match[0], ...parsed, inLi });
      lastIndex = paraRe.lastIndex;
    }
    const tail = html.slice(lastIndex);
    if (items.length === 0) return html;

    // Determine if we have at least a small run of list-like paragraphs
  const listLikeCount = items.filter(i => i.type !== null).length;
    if (listLikeCount < 2) return html;

    // Rebuild
    const out: string[] = [];
    let i = 0;
    while (i < items.length) {
      const it = items[i];
      out.push(paras[i] || '');
  if (it.type === null) {
        out.push(it.raw); // keep original paragraph
        i++;
        continue;
      }
      // Start a list block
      const stack: Array<{ type: 'ul'|'ol'; level: number }> = [];
      const closeUntil = (lvl: number) => {
        while (stack.length > 0 && stack[stack.length - 1].level >= lvl) {
          out.push(`</li>`);
          const last = stack.pop()!;
          out.push(`</${last.type}>`);
        }
      };
      const openLevel = (type: 'ul'|'ol', level: number) => {
        out.push(`<${type}>`);
        stack.push({ type, level });
      };
      // Open first level
      openLevel(it.type, it.level);
      out.push(`<li>${it.text}`);
      i++;
      while (i < items.length && items[i].type !== null) {
        const cur = items[i];
        // change of list type or level
        if (cur.level > stack[stack.length - 1].level) {
          // deeper
          openLevel(cur.type!, cur.level);
          out.push(`<li>${cur.text}`);
        } else if (cur.level === stack[stack.length - 1].level && cur.type === stack[stack.length - 1].type) {
          // same level
          out.push(`</li><li>${cur.text}`);
        } else {
          // shallower or type change
          closeUntil(cur.level);
          if (stack.length === 0 || stack[stack.length - 1].type !== cur.type) {
            openLevel(cur.type!, cur.level);
          }
          out.push(`<li>${cur.text}`);
        }
        i++;
      }
      // close all
      closeUntil(0);
    }
    out.push(tail);
    return out.join('');
  }

  private parseListishParagraph(attrs: string, content: string, inLi = false): { type: 'ul'|'ol'|null; level: number; text: string } {
    if (inLi) {
      return { type: null, level: 0, text: content };
    }
    const style = attrs.match(/style\s*=\s*"([^"]*)"/i)?.[1] || '';
    const marginPt = style.match(/margin-left\s*:\s*(\d+)pt/i);
    let level = marginPt ? Math.max(0, Math.floor(parseInt(marginPt[1], 10) / 18)) : 0;
    const leadingNbsp = (content.match(/^(?:&nbsp;|\s|\u00a0)+/) || [''])[0];
    if (leadingNbsp && !marginPt) {
      level = Math.floor((leadingNbsp.replace(/\s/g, '  ').length) / 4);
    }
    // Detect bullets / numbering
    const txt = content.replace(/<[^>]+>/g, '').trim();
    const bulletMatch = txt.match(/^([•◦▪\-\*\u2022\u25E6\u25AA])\s+(.*)$/);
    const numMatch = txt.match(/^((?:\d+|[ivxlcdm]+|[a-z])(?:[\.)]))\s+(.*)$/i);
    if (bulletMatch) {
      return { type: 'ul', level, text: content.replace(/^([\s\S]*?)([•◦▪\-\*\u2022\u25E6\u25AA])\s+/,'') };
    }
    if (numMatch) {
      // For ordered lists we leave numbering to Markdown; keep text only here
      const rest = content.replace(/^[\s\S]*?((?:\d+|[ivxlcdm]+|[a-z])[\.)])\s+/i, '');
      return { type: 'ol', level, text: rest };
    }
    return { type: null, level: 0, text: content };
  }

  /**
   * Convert HTML table to Markdown
   */
  // Removed DOM-dependent table conversion (rely on default Turndown behavior)

  /**
   * Detect code language from HTML element
   */
  private detectCodeLanguage(element: HTMLElement): string {
    const className = element.className || '';
    const content = element.textContent || '';

    // Check for language hints in class names
    const languagePatterns = [
      { pattern: /lang-(\w+)/, group: 1 },
      { pattern: /language-(\w+)/, group: 1 },
      { pattern: /(\w+)-code/, group: 1 },
    ];

    for (const { pattern, group } of languagePatterns) {
      const match = className.match(pattern);
      if (match) {
        return match[group];
      }
    }

    // Try to detect language from content
    if (content.includes('function') && content.includes('{')) {
      return 'javascript';
    }
    if (content.includes('def ') && content.includes(':')) {
      return 'python';
    }
    if (content.includes('public class') || content.includes('private ')) {
      return 'java';
    }
    if (content.includes('#include') || content.includes('int main')) {
      return 'c';
    }

    return '';
  }

  /**
   * Post-process markdown
   */
  private postProcessMarkdown(markdown: string, options: DocxToMarkdownOptions): string {
  let processed = markdown;

    // Normalize whitespace
    processed = StringUtils.normalizeWhitespace(processed);

    // Fix common formatting issues
    processed = this.fixMarkdownFormatting(processed);

  // Clean up excessive line breaks
    processed = processed.replace(/\n{3,}/g, '\n\n');

  // Unescape heading text like "# 1\. Intro" → "# 1. Intro"
  processed = this.unescapeHeadingDots(processed);

  // Optionally add anchors to headings to support internal links
  processed = this.applyHeadingAnchors(processed, options);

    // Attempt to fence obvious code blocks if not already fenced
    processed = this.fenceLikelyCodeBlocks(processed);

  // Normalize nested list indentation to 2 spaces per level
  processed = this.normalizeNestedLists(processed);

    // Trim leading/trailing whitespace
    processed = processed.trim();

    return processed;
  }

  private normalizeInternalAnchor(href: string, linkText?: string): string {
  // href like "#_8._Remarques" or "#Introduction" → "#8-remarques" or "#introduction"
  const fromText = (t: string) => `#${this.githubSlug(t.replace(/\\\./g, '.'))}`;
    if (linkText && linkText.trim()) {
      return fromText(linkText.trim());
    }
    const frag = href.replace(/^#/, '').replace(/^_+/, '');
    // Replace punctuation with spaces before slugging to avoid stray underscores
  const cleaned = decodeURIComponent(frag).replace(/[._]+/g, ' ').replace(/\s+/g, ' ').trim();
    return fromText(cleaned);
  }

  // Removed explicit heading anchor injection; we normalize links instead

  private fenceLikelyCodeBlocks(md: string): string {
    const lines = md.split(/\r?\n/);
    const out: string[] = [];
    let i = 0;
    while (i < lines.length) {
      // skip if already in fence
      if (/^\s*```/.test(lines[i])) {
        out.push(lines[i++]);
        while (i < lines.length && !/^\s*```/.test(lines[i])) out.push(lines[i++]);
        if (i < lines.length) out.push(lines[i++]);
        continue;
      }
      // detect start of code-ish block: not heading/list/blockquote/table row and contains code punctuation
      const isMarker = (s: string) => /^(#{1,6}\s|\s*[-*+]\s|\s*\d+\.\s|>\s|\|)/.test(s);
      const hasMarkdownLink = (s: string) => /!?\[[^\]]+\]\([^\)]+\)/.test(s);
      const looksCode = (s: string) => {
        if (hasMarkdownLink(s)) return false;
        return /[{;}=]|\b(def|function|class|console\.|import|return|let\s|const\s|var\s|#include)\b/.test(s) ||
               /^\s{4,}|\t/.test(s);
      };
      if (!isMarker(lines[i]) && looksCode(lines[i])) {
        const start = i;
        const block: string[] = [];
        while (i < lines.length && lines[i].trim() !== '' && !/^\s*```/.test(lines[i]) && !isMarker(lines[i])) {
          block.push(lines[i]);
          i++;
        }
        // Abort if the collected block contains Markdown links (likely prose)
        if (block.some(hasMarkdownLink)) {
          // not code; output original lines
          for (const l of block) out.push(l);
          continue;
        }
        if (block.length >= 2) {
          out.push('```');
          out.push(...block);
          out.push('```');
          continue;
        } else {
          // not enough lines to consider a block
          i = start;
        }
      }
      out.push(lines[i++]);
    }
    return out.join('\n');
  }

  /**
   * Fix common Markdown formatting issues
   */
  private fixMarkdownFormatting(markdown: string): string {
    let fixed = markdown;

    // Fix headings with extra spaces
    fixed = fixed.replace(/^(#{1,6})\s+(.+)$/gm, '$1 $2');

    // Fix list formatting
  // Preserve indentation for nested lists while normalizing bullet symbols
  fixed = fixed.replace(/^(\s*)([-*+])\s+/gm, '$1- ');
  fixed = fixed.replace(/^(\s*)(\d+)\.\s+/gm, '$1$2. ');

    // Fix blockquote formatting
    fixed = fixed.replace(/^>\s*/gm, '> ');

    // Fix code block formatting
    fixed = fixed.replace(/```\s*\n/g, '```\n');
    fixed = fixed.replace(/\n\s*```/g, '\n```');

    return fixed;
  }

  private unescapeHeadingDots(md: string): string {
    const lines = md.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      if (/^#{1,6}\s+/.test(lines[i])) {
        lines[i] = lines[i].replace(/\\\./g, '.');
      }
    }
    return lines.join('\n');
  }

  private normalizeNestedLists(md: string): string {
    const lines = md.split(/\r?\n/);
    const bulletRe = /^(\s*)([-*+])\s+/;
    const numRe = /^(\s*)(\d+)\.\s+/;
    let lastListIndent = 0;
    let lastWasList = false;
    for (let i = 0; i < lines.length; i++) {
      const l = lines[i];
      const mBullet = l.match(bulletRe);
      const mNum = !mBullet && l.match(numRe);
      if (mBullet || mNum) {
        const ws = (mBullet ? mBullet[1] : (mNum ? mNum[1] : '')) || '';
        // Normalize tabs to two spaces
        let spaces = ws.replace(/\t/g, '  ').length;
        // Ensure increments in steps of 2 when nesting
        if (lastWasList && spaces > lastListIndent && spaces < lastListIndent + 2) {
          spaces = lastListIndent + 2;
        }
        // Round to even number of spaces
        if (spaces % 2 === 1) spaces += 1;
        const newWs = ' '.repeat(spaces);
        if (mBullet) {
          lines[i] = l.replace(bulletRe, `${newWs}- `);
        } else if (mNum) {
          lines[i] = l.replace(numRe, `${newWs}${mNum[2]}. `);
        }
        lastWasList = true;
        lastListIndent = spaces;
      } else if (l.trim() === '') {
        // reset between list blocks
        lastWasList = false;
        lastListIndent = 0;
      } else {
        lastWasList = false;
      }
    }
    return lines.join('\n');
  }

  private applyHeadingAnchors(md: string, options: DocxToMarkdownOptions): string {
  const mode = options.headingAnchors || 'none';
    if (mode === 'none') return md;
    const lines = md.split(/\r?\n/);
    const out: string[] = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const m = line.match(/^(#{1,6})\s+(.*)$/);
      if (!m) { out.push(line); continue; }
      const text = m[2].trim();
      const slug = this.githubSlug(text);
      if (mode === 'html') {
        out.push(`<a id="${slug}"></a>`);
        out.push(line);
      } else if (mode === 'pandoc') {
        out.push(`${line} {#${slug}}`);
      } else {
        out.push(line);
      }
    }
    return out.join('\n');
  }

  // GitHub-style slug (close enough): lowercase, trim, remove accents, replace spaces with '-', remove invalids
  private githubSlug(text: string): string {
    const from = 'àáâäæãåāçćčđèéêëēėęîïíīįìłñńôöòóõøōßśšûüùúūýÿžźż·/_,:;'
    const to   = 'aaaaaaaaacccdeeeeeeeiiiiii lnnooooooos suuuuuyyz zz------'
      .replace(/\s+/g, '');
    const p = new RegExp(from.split('').join('|'), 'g');
    return text.toLowerCase()
      .replace(p, c => to[from.indexOf(c)] || '')
      .replace(/&amp;|&/g, 'and')
      .replace(/\./g, '')
      .replace(/[^a-z0-9\s-]/g, '')
      .trim()
      .replace(/[\s_-]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }

  /**
   * Save extracted images to filesystem
   */
  private async saveExtractedImages(outputDir: string): Promise<void> {
    if (this.extractedImages.length === 0) {
      return;
    }

    await FileUtils.ensureWritableDir(outputDir);

    for (const image of this.extractedImages) {
      const imagePath = path.join(outputDir, image.filename);
      await fs.writeFile(imagePath, image.buffer);
      logger.debug(`Saved extracted image: ${imagePath}`);
    }

    logger.info(`Saved ${this.extractedImages.length} extracted images to ${outputDir}`);
  }

  /**
   * Get conversion warnings
   */
  private getConversionWarnings(messages: any[], options?: DocxToMarkdownOptions): string[] {
    const warnings: string[] = [];

    const messagesByType = messages.reduce((acc, msg) => {
      acc[msg.type] = (acc[msg.type] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    if (messagesByType.warning > 0) {
      warnings.push(`${messagesByType.warning} formatting warnings during conversion`);
      if (options?.detailedWarnings) {
        messages.filter(m => m.type === 'warning').slice(0, 50).forEach((m: any, idx: number) => {
          const detail = [m.message, m.error?.message].filter(Boolean).join(' - ');
          const where = m.path ? ` @ ${m.path}` : '';
          warnings.push(`  ${idx + 1}. ${detail}${where}`);
        });
      }
    }

    if (messagesByType.error > 0) {
      warnings.push(`${messagesByType.error} errors during conversion`);
      if (options?.detailedWarnings) {
        messages.filter(m => m.type === 'error').slice(0, 50).forEach((m: any, idx: number) => {
          const detail = [m.message, m.error?.message].filter(Boolean).join(' - ');
          const where = m.path ? ` @ ${m.path}` : '';
          warnings.push(`  ${idx + 1}. ${detail}${where}`);
        });
      }
    }

    if (this.extractedImages.length > 0) {
      warnings.push(`${this.extractedImages.length} images extracted`);
    }

    return warnings;
  }

  /**
   * Cleanup resources
   */
  private cleanup(): void {
    this.extractedImages = [];
  }
}
