import * as mammoth from 'mammoth';
import TurndownService from 'turndown';
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

      // Convert HTML to Markdown
      const { result: markdown, elapsed: conversionTime } = 
        await PerformanceUtils.measureAsync(async () => {
          return this.convertHtmlToMarkdown(extraction.html, options);
        }, 'html-to-markdown');

      logger.debug(`HTML to Markdown conversion completed in ${conversionTime}ms`);

      // Save extracted images if requested
      if (options.extractImages && options.imageOutputDir) {
        await this.saveExtractedImages(options.imageOutputDir);
      }

      // Post-process markdown
      const processedMarkdown = this.postProcessMarkdown(markdown, options);

      const totalTime = Date.now() - startTime;

      const metadata = {
        inputSize: docxBuffer.length,
        outputSize: Buffer.from(processedMarkdown, 'utf-8').length,
        processingTime: totalTime,
        imageCount: this.extractedImages.length,
      };

      logger.info('DOCX to Markdown conversion completed successfully', metadata);

      return {
        success: true,
        output: processedMarkdown,
        metadata,
        warnings: this.getConversionWarnings(extraction.messages),
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

    // Custom rules for better formatting
    this.addCustomTurndownRules(turndown, options);

    return turndown;
  }

  /**
   * Add custom Turndown rules
   */
  private addCustomTurndownRules(turndown: TurndownService, options: DocxToMarkdownOptions): void {
    // Preserve code blocks
    turndown.addRule('codeBlock', {
      filter: ['pre'],
      replacement: (content, node) => {
        const language = this.detectCodeLanguage(node as HTMLElement);
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

    // Handle tables with better formatting
    turndown.addRule('table', {
      filter: 'table',
      replacement: (content, node) => {
        return this.convertTableToMarkdown(node as HTMLTableElement);
      },
    });

    // Handle images
    turndown.addRule('image', {
      filter: 'img',
      replacement: (content, node) => {
        const img = node as HTMLImageElement;
        const alt = img.alt || 'Image';
        const src = img.src || '';
        
        // Check if this is an extracted image
        const extractedImage = this.extractedImages.find(img => src.includes(img.filename));
        const imagePath = extractedImage 
          ? `./${options.imageOutputDir || 'images'}/${extractedImage.filename}`
          : src;

        return `![${alt}](${imagePath})`;
      },
    });

    // Handle hyperlinks
    turndown.addRule('hyperlink', {
      filter: (node) => {
        return node.nodeName === 'A' && (node as HTMLAnchorElement).href ? true : false;
      },
      replacement: (content, node) => {
        const link = node as HTMLAnchorElement;
        const href = link.href;
        
        // Handle internal links/bookmarks
        if (href.startsWith('#')) {
          return `[${content}](${href})`;
        }
        
        return `[${content}](${href})`;
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
      
      // Lists
      "p[style-name='List Paragraph'] => p:fresh",
      
      // Normal paragraphs
      "p[style-name='Normal'] => p:fresh",
    ];

    return {
      styleMap,
      includeDefaultStyleMap: true,
      includeEmbeddedStyleMap: true,
      convertImage: (image: any) => this.handleImageExtraction(image, options),
      ignoreEmptyParagraphs: false,
      preserveLineBreaks: options.preserveFormatting || false,
    };
  }

  /**
   * Extract DOCX content
   */
  private async extractDocxContent(
    docxBuffer: Buffer,
    mammothOptions: any,
    options: DocxToMarkdownOptions
  ): Promise<{ html: string; messages: any[] }> {
    try {
      const result = await mammoth.convertToHtml(docxBuffer as any, mammothOptions);
      
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
      logger.error('Failed to extract DOCX content', { error });
      throw error;
    }
  }

  /**
   * Handle image extraction
   */
  private handleImageExtraction(image: any, options: DocxToMarkdownOptions): any {
    if (!options.extractImages) {
      return image;
    }

    try {
      const extension = this.getImageExtension(image.contentType);
      const filename = `image_${this.extractedImages.length + 1}${extension}`;
      
      this.extractedImages.push({
        filename,
        buffer: image.buffer,
      });

      logger.debug(`Extracted image: ${filename}`);
      
      return {
        src: filename,
        altText: image.altText || 'Extracted image',
      };
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

    // Normalize whitespace
    cleaned = cleaned.replace(/\s+/g, ' ');
    cleaned = cleaned.replace(/>\s+</g, '><');

    // Fix list formatting
    cleaned = this.fixListFormatting(cleaned);

    // Fix table formatting
    cleaned = this.fixTableFormatting(cleaned);

    return cleaned;
  }

  /**
   * Fix list formatting in HTML
   */
  private fixListFormatting(html: string): string {
    // Convert Word list paragraphs to proper HTML lists
    return html.replace(
      /<p[^>]*>\s*([?????\-\*]|\d+\.)\s*([^<]+)<\/p>/g,
      '<li>$2</li>'
    );
  }

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
    
    return fixed;
  }

  /**
   * Convert HTML table to Markdown
   */
  private convertTableToMarkdown(table: HTMLTableElement): string {
    const rows: string[][] = [];
    
    // Extract table data
    const tableRows = table.querySelectorAll('tr');
    for (let i = 0; i < tableRows.length; i++) {
      const row = tableRows[i];
      const cells: string[] = [];
      const tableCells = row.querySelectorAll('td, th');
      
      for (let j = 0; j < tableCells.length; j++) {
        const cell = tableCells[j];
        cells.push(cell.textContent?.trim() || '');
      }
      
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    if (rows.length === 0) {
      return '';
    }

    // Convert to Markdown table
    const markdownRows: string[] = [];
    const maxCols = Math.max(...rows.map(row => row.length));

    // Header row
    if (rows.length > 0) {
      const headerRow = rows[0].map(cell => cell || '').slice(0, maxCols);
      while (headerRow.length < maxCols) {
        headerRow.push('');
      }
      markdownRows.push(`| ${headerRow.join(' | ')} |`);
      
      // Separator row
      const separator = Array(maxCols).fill('---').join(' | ');
      markdownRows.push(`| ${separator} |`);
    }

    // Data rows
    for (let i = 1; i < rows.length; i++) {
      const dataRow = rows[i].map(cell => cell || '').slice(0, maxCols);
      while (dataRow.length < maxCols) {
        dataRow.push('');
      }
      markdownRows.push(`| ${dataRow.join(' | ')} |`);
    }

    return '\n' + markdownRows.join('\n') + '\n';
  }

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

    // Trim leading/trailing whitespace
    processed = processed.trim();

    return processed;
  }

  /**
   * Fix common Markdown formatting issues
   */
  private fixMarkdownFormatting(markdown: string): string {
    let fixed = markdown;

    // Fix headings with extra spaces
    fixed = fixed.replace(/^(#{1,6})\s+(.+)$/gm, '$1 $2');

    // Fix list formatting
    fixed = fixed.replace(/^[\s]*([?????\-\*])\s+/gm, '- ');
    fixed = fixed.replace(/^[\s]*(\d+)\.\s+/gm, '$1. ');

    // Fix blockquote formatting
    fixed = fixed.replace(/^>\s*/gm, '> ');

    // Fix code block formatting
    fixed = fixed.replace(/```\s*\n/g, '```\n');
    fixed = fixed.replace(/\n\s*```/g, '\n```');

    return fixed;
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
  private getConversionWarnings(messages: any[]): string[] {
    const warnings: string[] = [];

    const messagesByType = messages.reduce((acc, msg) => {
      acc[msg.type] = (acc[msg.type] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    if (messagesByType.warning > 0) {
      warnings.push(`${messagesByType.warning} formatting warnings during conversion`);
    }

    if (messagesByType.error > 0) {
      warnings.push(`${messagesByType.error} errors during conversion`);
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
