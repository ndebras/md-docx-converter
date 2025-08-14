import {
  Document,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ImageRun,
  ExternalHyperlink,
  InternalHyperlink,
  BookmarkStart,
  BookmarkEnd,
  TableOfContents,
  SectionType,
  PageBreak,
  Packer,
  LevelFormat,
} from 'docx';
import { marked } from 'marked';
import { MermaidPNGProcessor } from '../processors/mermaid-png-processor';
import { LinkProcessor } from '../processors/link-processor';
import { DocumentTemplates } from '../styles/document-templates';
import { Logger } from '../utils/logger';
import type { 
  MarkdownToDocxOptions, 
  ConversionResult, 
  ConversionMetadata, 
  DocumentTemplate 
} from '../types';
import sharp from 'sharp';

export class MarkdownToDocxConverter {
  private mermaidProcessor: MermaidPNGProcessor;
  private linkProcessor: LinkProcessor;
  private logger: Logger;
  private bookmarkCounter = 0;
  private headings: Array<{ level: number; text: string; anchor: string }> = [];
  private mermaidImages: Map<string, Buffer> = new Map();
  // Track ordered list numbering references so each new list restarts at 1
  private orderedListCounter = 0;
  private usedOrderedListRefs: Set<string> = new Set();
  private template: any; // Store the current template

  constructor() {
    this.mermaidProcessor = new MermaidPNGProcessor();
    this.linkProcessor = new LinkProcessor();
    this.logger = new Logger({ level: 'info' });
  }

  /**
   * Convert a Markdown string into a DOCX buffer using docx.
   * Pipeline:
   * 1) Strip front matter and merge metadata
   * 2) Pre-process Mermaid (generate PNGs)
   * 3) Normalize malformed nested links (e.g., [A]([B](URL)))
   * 4) Lex markdown via marked and build a docx AST with processTokens
   * 5) Optionally inject a TOC, then pack to a DOCX buffer
   */
  async convert(
    markdownContent: string,
    options: MarkdownToDocxOptions = {}
  ): Promise<ConversionResult> {
    try {
      this.logger.info('Starting Markdown to DOCX conversion');
      const startTime = Date.now();

  // Reset per-conversion state
  this.orderedListCounter = 0;
  this.usedOrderedListRefs.clear();

      // Guard: empty markdown should fail fast with a clear error
      if (!markdownContent || markdownContent.trim().length === 0) {
        return {
          success: false,
          error: {
            code: 'EMPTY_INPUT',
            message: 'Empty markdown content',
          }
        };
      }

      // Remove optional YAML front matter and extract simple metadata
      const { content: contentNoFM, meta } = this.stripFrontMatter(markdownContent);

      // Merge parsed metadata into options (caller options win)
      const effOptions: MarkdownToDocxOptions = {
        ...options,
        title: options.title || meta.title,
        author: options.author || meta.author,
        description: options.description || meta.description,
        subject: options.subject || meta.description,
      };

      // Process Mermaid diagrams first and get PNG images
      const mermaidResult = await this.mermaidProcessor.processContent(contentNoFM);
      
      // Store images for later use in DOCX
      mermaidResult.images.forEach(image => {
        this.mermaidImages.set(image.id, image.buffer);
      });

  // Normalize malformed nested links like [A]([B](URL)) -> [A](URL)
  const normalizedContent = this.normalizeNestedLinks(mermaidResult.content);
      
  // Parse markdown
  const tokens = marked.lexer(normalizedContent);
      
      // Process links
      const linkResult = options.preserveLinks 
        ? this.linkProcessor.processContent(normalizedContent)
        : { content: normalizedContent, links: [], headings: [] };

      // Extract headings for TOC
      this.headings = this.extractHeadings(tokens);

      // Get template and store it for use in element creation methods
      this.template = DocumentTemplates.getTemplate(effOptions.template || 'professional-report');

      // Convert to DOCX
      const elements = await this.processTokens(tokens);
      
      // Add Table of Contents if requested
      if (options.tocGeneration) {
        const tocParagraphs = this.createTableOfContentsWithLinks();
        elements.unshift(...tocParagraphs);
      }

      // Create document
  const document = this.createDocument(elements, effOptions);
      
      // Generate buffer
      const buffer = await Packer.toBuffer(document);

      const processingTime = Date.now() - startTime;
      
      const metadata: ConversionMetadata = {
        inputSize: markdownContent.length,
        outputSize: buffer.length,
        processingTime,
        pageCount: Math.ceil(buffer.length / 4000), // Rough estimation
        mermaidDiagramCount: mermaidResult.diagramCount || 0,
        internalLinkCount: linkResult.links.filter((l: any) => l.type === 'internal').length,
        externalLinkCount: linkResult.links.filter((l: any) => l.type === 'external').length,
      };

      this.logger.info(`Conversion completed in ${processingTime}ms`);

      // Nettoyer les ressources du processeur Mermaid
      await this.mermaidProcessor.cleanup();
      this.mermaidImages.clear();

      return {
        success: true,
        output: buffer,
        metadata,
        warnings: []
      };

    } catch (error) {
      this.logger.error('Conversion failed:', { error: error instanceof Error ? error.message : String(error) });
      
      // Nettoyer en cas d'erreur aussi
      await this.mermaidProcessor.cleanup();
      this.mermaidImages.clear();
      
        return {
        success: false,
        error: {
          message: error instanceof Error ? error.message : 'Unknown error',
          code: 'CONVERSION_FAILED',
          details: error instanceof Error ? { stack: error.stack } : { error: String(error) }
        }
      };
    }
  }

  /**
   * Strip YAML front matter at the beginning of the markdown file and parse a few common keys.
   */
  private stripFrontMatter(markdown: string): { content: string; meta: { title?: string; author?: string; description?: string } } {
    const meta: { title?: string; author?: string; description?: string } = {};
    if (!markdown) return { content: markdown, meta };

    // Only treat as front matter if it starts with --- on the very first line
    if (!markdown.startsWith('---')) return { content: markdown, meta };

    // Find the closing --- after the first line
    const end = markdown.indexOf('\n---', 3);
    if (end === -1) return { content: markdown, meta };

    const fmBlock = markdown.substring(3, end).trim();
    const rest = markdown.substring(end + 4); // skip leading \n before second ---

    // Parse simple key: value pairs (quoted or not)
    const lines = fmBlock.split(/\r?\n/);
    for (const line of lines) {
      const m = line.match(/^\s*([A-Za-z0-9_-]+)\s*:\s*(.*)\s*$/);
      if (!m) continue;
      const key = m[1].toLowerCase();
      let value = m[2];
      // Trim surrounding quotes if present
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith('\'') && value.endsWith('\''))) {
        value = value.substring(1, value.length - 1);
      }
      if (['title','author','description','subject'].includes(key)) {
        (meta as any)[key] = value;
      }
    }

    return { content: rest.replace(/^\s*\r?\n/, ''), meta };
  }

  /**
   * Walk the list of top-level marked tokens and convert each into
   * a docx Paragraph/Table (or a list of paragraphs for lists).
   */
  private async processTokens(tokens: any[]): Promise<(Paragraph | Table)[]> {
    const elements: (Paragraph | Table)[] = [];

    for (const token of tokens) {
      const result = await this.processToken(token);
      if (result) {
        if (Array.isArray(result)) {
          elements.push(...result);
        } else {
          elements.push(result);
        }
      }
    }

    return elements;
  }

  /**
   * Convert a single marked token to docx structures. Inline content is
   * delegated to processInlineTokens/WithStyle.
   */
  private async processToken(token: any): Promise<Paragraph | Table | Paragraph[] | null> {
    switch (token.type) {
      case 'heading':
        return this.createHeading(token);
      case 'paragraph':
        return await this.createParagraph(token);
      case 'list':
        return this.createList(token);
      case 'blockquote':
        return this.createBlockquote(token);
      case 'code':
        return this.createCodeBlock(token);
      case 'table':
        return this.createTable(token);
      case 'hr':
        return this.createHorizontalRule();
      case 'html':
        return this.createHtml(token);
      case 'image':
        return await this.createImageParagraph(token);
      case 'space':
        return null;
      default:
        this.logger.warn(`Unsupported token type: ${token.type}`);
        return null;
    }
  }

  private createHeading(token: any): Paragraph {
    const anchor = this.sanitizeBookmarkName(this.generateAnchor(token.text));
    this.bookmarkCounter += 1;
    const headingText = this.decodeHtmlEntities(token.text);
    
    // Get the appropriate heading style from template
    const headingStyleKey = `heading${token.depth}`;
    const headingStyle = this.template?.styles?.headings?.[headingStyleKey];

    return new Paragraph({
      heading: this.getHeadingLevel(token.depth),
      children: [
        // Use anchor as bookmark name so InternalHyperlink can target it
        new BookmarkStart(anchor, this.bookmarkCounter),
        new TextRun({
          text: headingText,
          bold: headingStyle?.run?.bold !== undefined ? headingStyle.run.bold : true,
          size: headingStyle?.run?.size || 24,
          color: headingStyle?.run?.color || '000000',
          font: headingStyle?.run?.font || 'Calibri',
        }),
        new BookmarkEnd(this.bookmarkCounter),
      ],
      spacing: headingStyle?.paragraph?.spacing,
      border: headingStyle?.paragraph?.border,
    });
  }

  private async createParagraph(token: any): Promise<Paragraph> {
    // Check if this paragraph contains only an image
    if (token.tokens && token.tokens.length === 1 && token.tokens[0].type === 'image') {
      return await this.createImageParagraph(token.tokens[0]);
    }
    
    // Get paragraph style from template
    const defaultDoc = this.template?.styles?.default?.document;
    
    return new Paragraph({
      children: this.processInlineTokens(token.tokens || [{ type: 'text', text: token.text }]),
      spacing: defaultDoc?.paragraph?.spacing,
      alignment: defaultDoc?.paragraph?.alignment,
    });
  }

  private async createImageParagraph(token: any): Promise<Paragraph> {
    // Check if it's a Mermaid diagram reference
    if (token.href && token.href.includes('mermaid-') && token.href.endsWith('.png')) {
      const mermaidId = token.href.replace('mermaid-', '').replace('.png', '');
      const imageBuffer = this.mermaidImages.get(mermaidId);
      if (imageBuffer) {
        // Determine intrinsic size using sharp metadata; fallback to defaults
        let width = 600;
        let height = 450;
        try {
          const meta = await sharp(imageBuffer).metadata();
          if (meta.width && meta.height) {
            width = meta.width;
            height = meta.height;
          }
        } catch {
          // ignore
        }
        return new Paragraph({
          children: [
            new ImageRun({
              data: imageBuffer,
              transformation: { width, height },
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 200 },
        });
      }
    }
    // Fallback for other images or missing Mermaid images
    return new Paragraph({
      children: [
        new TextRun({
          text: `[Image: ${token.alt || token.href || 'Non disponible'}]`,
          italics: true,
          color: '0066CC',
          size: 24,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 200 },
    });
  }

  private createList(token: any, level: number = 0, orderedRef?: string): Paragraph[] {
    const paragraphs: Paragraph[] = [];

    // Determine numbering reference for this list (ordered lists get a fresh reference per list)
    const currentOrderedRef = token.ordered
      ? (orderedRef || this.newOrderedListReference())
      : undefined;

    token.items.forEach((item: any, index: number) => {
      // Déterminer si c'est une tâche (checkbox) et choisir le préfixe
      // Detect GitHub task list markers
      let isTask = !!item.task || /^\s*\[(x|X| )\]/.test(item.raw || '');
      let isChecked = !!item.checked;

      // Les items de liste de marked ont généralement des tokens inline (texte, strong, em, link, code...)
      const inlineTokens = Array.isArray(item.tokens)
        ? item.tokens.filter((t: any) => t.type !== 'list')
        : [{ type: 'text', text: typeof item.text === 'string' ? item.text : String(item.text || '') }];

      // Fallback detection based on first inline text content
      if (!isTask && inlineTokens.length > 0) {
        const firstText = (inlineTokens[0].type === 'text') ? inlineTokens[0].text
          : (inlineTokens[0].type === 'paragraph' && Array.isArray(inlineTokens[0].tokens) && inlineTokens[0].tokens[0]?.type === 'text')
            ? inlineTokens[0].tokens[0].text
            : '';
        const m = /^\s*\[(x|X| )\]\s*/.exec(firstText || '');
        if (m) {
          isTask = true;
          isChecked = m[1].toLowerCase() === 'x';
        }
      }

  const checkbox = isTask ? (isChecked ? '☑' : '☐') : null;

      // Si item de tâche, enlever le marqueur [ ] / [x] du premier token texte
      if (isTask && inlineTokens.length > 0) {
        if (inlineTokens[0].type === 'text' && typeof inlineTokens[0].text === 'string') {
          inlineTokens[0].text = inlineTokens[0].text.replace(/^\s*\[(x|X| )\]\s*/, '');
        } else if (inlineTokens[0].type === 'paragraph' && Array.isArray(inlineTokens[0].tokens) && inlineTokens[0].tokens.length > 0) {
          const first = inlineTokens[0].tokens[0];
          if (first && first.type === 'text' && typeof first.text === 'string') {
            first.text = first.text.replace(/^\s*\[(x|X| )\]\s*/, '');
          }
        }
      }

      const formattedRuns = this.processInlineTokens(inlineTokens);

      const listChildren: any[] = [];
      if (isTask) {
        const defaultStyle = this.getDefaultTextStyle();
        listChildren.push(new TextRun({ text: `${checkbox} `, ...defaultStyle }));
      }
      listChildren.push(...formattedRuns);

      const paraOptions: any = { children: listChildren };
      if (!isTask) {
        paraOptions.numbering = {
          reference: token.ordered ? (currentOrderedRef as string) : 'md-bullet',
          level: Math.max(0, Math.min(2, level)),
        };
      } else {
        // Visual indent for task items to align with list levels
        paraOptions.indent = { left: 720 * (level + 1) };
      }

      paragraphs.push(new Paragraph(paraOptions));

      // Gérer d'éventuelles sous-listes imbriquées
      if (Array.isArray(item.tokens)) {
        item.tokens
          .filter((t: any) => t.type === 'list')
          .forEach((subList: any) => {
            // Nested ordered lists should restart at 1 → use a fresh reference per sublist
            const subParas = this.createList(subList, level + 1, undefined);
            paragraphs.push(...subParas);
          });
      }
    });

    return paragraphs;
  }

  private createBlockquote(token: any): Paragraph[] {
    const paragraphs: Paragraph[] = [];
    
    if (token.tokens) {
      token.tokens.forEach((subToken: any) => {
        if (subToken.type === 'paragraph') {
          paragraphs.push(new Paragraph({
            children: this.processInlineTokens(subToken.tokens || [{ type: 'text', text: subToken.text }]),
            indent: {
              left: 720, // 0.5 inch
            },
            border: {
              left: {
                color: 'CCCCCC',
                size: 6,
                style: 'single',
              },
            },
          }));
        }
      });
    }

    return paragraphs;
  }

  private createCodeBlock(token: any): Paragraph[] {
    // Split the code into lines and create a separate paragraph for each line
    const lines = token.text.split('\n');
    const paragraphs: Paragraph[] = [];
    
    lines.forEach((line: string, index: number) => {
      paragraphs.push(new Paragraph({
        children: [
          new TextRun({
            text: line || ' ', // Use space for empty lines to preserve line breaks
            font: 'Courier New',
            size: 18,
            color: '333333',
          })
        ],
        shading: {
          fill: 'F8F8F8',
        },
        border: {
          top: { style: 'single', size: 1, color: 'DDDDDD' },
          bottom: { style: 'single', size: 1, color: 'DDDDDD' },
          left: { style: 'single', size: 1, color: 'DDDDDD' },
          right: { style: 'single', size: 1, color: 'DDDDDD' },
        },
        spacing: {
          // Only add spacing before for the first line and after for the last line
          before: index === 0 ? 200 : 0,
          after: index === lines.length - 1 ? 200 : 0,
        },
      }));
    });

    return paragraphs;
  }

  private createTable(token: any): Table {
    const cellRuns = (cell: any, isHeader = false): TextRun[] => {
      // Get table cell style from template
      const tableCellStyle = this.template?.styles?.table?.cell?.run;
      const defaultStyle = tableCellStyle ? {
        font: tableCellStyle.font,
        size: tableCellStyle.size,
        color: tableCellStyle.color
      } : this.getDefaultTextStyle();
      
      const baseStyle = isHeader ? { ...defaultStyle, bold: true } : defaultStyle;
      
      if (typeof cell === 'string') {
        return this.parseSimpleMarkdown(cell, baseStyle);
      }
      if (cell && Array.isArray(cell.tokens)) {
        return this.processInlineTokensWithStyle(cell.tokens, baseStyle);
      }
      if (cell && typeof cell.text === 'string') {
        return this.parseSimpleMarkdown(cell.text, baseStyle);
      }
      // Fallback: convertir en chaîne proprement
      return [new TextRun({ text: String(cell ?? ''), ...baseStyle })];
    };

    // Get table cell paragraph style from template
    const tableCellParagraphStyle = this.template?.styles?.table?.cell?.paragraph;

    const headerCells = token.header.map((cell: any) => new TableCell({
      children: [new Paragraph({ 
        children: cellRuns(cell, true),
        spacing: tableCellParagraphStyle?.spacing
      })],
      shading: { fill: 'E6E6E6' },
    }));

    const rows: TableRow[] = [new TableRow({ children: headerCells })];

    token.rows.forEach((rowData: any[]) => {
      const dataCells = rowData.map((cell: any) => new TableCell({
        children: [new Paragraph({ 
          children: cellRuns(cell, false),
          spacing: tableCellParagraphStyle?.spacing
        })],
      }));
      rows.push(new TableRow({ children: dataCells }));
    });

    return new Table({
      rows,
      width: { size: 100, type: WidthType.PERCENTAGE },
    });
  }

  /**
   * Basic HTML block handling. Mermaid blocks are handled upstream.
   * For other HTML, we strip tags and keep text to avoid raw HTML leakage.
   */
  private createHtml(token: any): Paragraph | null {
    // Handle Mermaid diagrams
    if (token.text.includes('mermaid')) {
      // This should be handled by the MermaidProcessor
      return null;
    }

    // Basic HTML handling
    const text = token.text.replace(/<[^>]*>/g, '');
    return new Paragraph({
      children: [new TextRun({ text })],
    });
  }

  private createHorizontalRule(): Paragraph {
    const defaultStyle = this.getDefaultTextStyle();
    return new Paragraph({
      children: [new TextRun({ text: '', ...defaultStyle })],
      border: {
        bottom: {
          color: 'CCCCCC',
          size: 6,
          style: 'single',
        },
      },
      spacing: {
        before: 200,
        after: 200,
      },
    });
  }

  // Fix malformed nested links like [Text]([InnerText](URL)) -> [Text](URL)
  private normalizeNestedLinks(md: string): string {
    if (!md) return md;
    // Repeat replacement a few times to catch multiple occurrences
    let out = md;
    const pattern = /\[([^\]]+)\]\(\s*\[([^\]]+)\]\(([^)]+)\)\s*\)/g; // [A]([B](URL))
    for (let i = 0; i < 5; i++) {
      const before = out;
      out = out.replace(pattern, '[$1]($3)');
      if (out === before) break;
    }
    return out;
  }

  private processInlineTokens(tokens: any[]): any[] {
    const runs: any[] = [];
    const defaultStyle = this.getDefaultTextStyle();

    // Use index-based loop to allow lookahead for HTML open/close tags
    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i];
      switch (token.type) {
        case 'paragraph':
          if (Array.isArray(token.tokens)) {
            runs.push(...this.processInlineTokens(token.tokens));
          } else if (token.text) {
            runs.push(...this.processTextWithInlineHtml(token.text));
          }
          break;
        case 'text':
          if (Array.isArray((token as any).tokens) && (token as any).tokens.length > 0) {
            runs.push(...this.processInlineTokens((token as any).tokens));
          } else {
            runs.push(...this.processTextWithInlineHtml(token.text));
          }
          break;
        case 'html': {
          const raw = String(token.text || '').trim().toLowerCase();
          // Standalone <br> should create a line break
          if (raw === '<br>' || raw === '<br/>' || raw === '<br />') {
            runs.push(new TextRun({ break: 1 } as any));
            break;
          }
          // Try to group <u>/<b>/<strong>/<i>/<em>/<s>/<del> blocks
          const tagToStyle: Record<string, any> = {
            '<u>': { underline: {} },
            '<b>': { bold: true },
            '<strong>': { bold: true },
            '<i>': { italics: true },
            '<em>': { italics: true },
            '<s>': { strike: true },
            '<del>': { strike: true },
          };
          const closeFor: Record<string, string> = {
            '<u>': '</u>',
            '<b>': '</b>',
            '<strong>': '</strong>',
            '<i>': '</i>',
            '<em>': '</em>',
            '<s>': '</s>',
            '<del>': '</del>',
          };
          if (raw in tagToStyle) {
            const closeTag = closeFor[raw];
            // Find matching closing tag ahead
            let j = i + 1;
            while (j < tokens.length) {
              const t = tokens[j];
              if (t.type === 'html' && String(t.text || '').trim().toLowerCase() === closeTag) break;
              j++;
            }
            if (j < tokens.length) {
              // Process inner tokens with style
              const inner = tokens.slice(i + 1, j);
              runs.push(...this.processInlineTokensWithStyle(inner, tagToStyle[raw]));
              i = j; // skip to closing
              break;
            }
            // If no closing found, ignore the tag (don’t render raw)
            break;
          }
          // If html token contains inline HTML with content (e.g., <u>text</u>)
          if (/[<>].*[<>]/.test(raw)) {
            runs.push(...this.processTextWithInlineHtml(token.text));
          }
          // Otherwise ignore standalone raw HTML
          break;
        }
        case 'strong':
          if (token.tokens && token.tokens.length > 0) {
            runs.push(...this.processInlineTokensWithStyle(token.tokens, { bold: true }));
          } else {
            runs.push(...this.processTextWithInlineHtml(token.text, { bold: true }));
          }
          break;
        case 'em':
          if (token.tokens && token.tokens.length > 0) {
            runs.push(...this.processInlineTokensWithStyle(token.tokens, { italics: true }));
          } else {
            runs.push(...this.processTextWithInlineHtml(token.text, { italics: true }));
          }
          break;
        case 'del':
        case 's':
          if (token.tokens && token.tokens.length > 0) {
            runs.push(...this.processInlineTokensWithStyle(token.tokens, { strike: true }));
          } else {
            runs.push(...this.processTextWithInlineHtml(token.text, { strike: true }));
          }
          break;
        case 'codespan':
        case 'code':
          runs.push(new TextRun({ text: this.decodeHtmlEntities(token.text), ...defaultStyle, font: 'Courier New' }));
          break;
        case 'link': {
          const rawHref = typeof token.href === 'string' ? token.href.trim() : '';
          let finalHref = rawHref;
          let finalText = this.decodeHtmlEntities(token.text || '');
          const nested = rawHref.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
          if (nested) {
            // Keep the outer link text, but use the inner URL
            finalHref = nested[2].trim();
          }
          const isExternal = /^(https?:\/\/|mailto:|ftp:\/\/)/i.test(finalHref);
          if (isExternal) {
            runs.push(new ExternalHyperlink({
              link: finalHref,
              children: [new TextRun({ text: finalText, ...defaultStyle, style: 'Hyperlink' })],
            }));
          } else {
            const anchorText = finalHref.replace(/^#/, '');
            const anchor = this.sanitizeBookmarkName(anchorText);
            runs.push(new InternalHyperlink({
              anchor,
              children: [new TextRun({ text: finalText, ...defaultStyle, color: '0000FF', underline: {} })],
            }));
          }
          break;
        }
        case 'image':
          if (token.href && token.href.startsWith('data:image/')) {
            runs.push(new TextRun({ text: `[${token.alt || 'Image'}]`, ...defaultStyle, italics: true, color: '666666' }));
          } else {
            runs.push(new TextRun({ text: `[Image: ${token.alt || token.href}]`, ...defaultStyle, italics: true, color: '666666' }));
          }
          break;
        default:
          if (token.text) {
            runs.push(...this.processTextWithInlineHtml(token.text));
          }
      }
    }

    return runs;
  }

  private processInlineTokensWithStyle(tokens: any[], baseStyle: any): any[] {
    const runs: any[] = [];
    const defaultStyle = this.getDefaultTextStyle();

    tokens.forEach(token => {
      switch (token.type) {
        case 'paragraph':
          if (Array.isArray(token.tokens)) {
            runs.push(...this.processInlineTokensWithStyle(token.tokens, { ...baseStyle }));
          } else if (token.text) {
            runs.push(...this.processTextWithInlineHtml(token.text, baseStyle));
          }
          break;
        case 'text':
          if (Array.isArray((token as any).tokens) && (token as any).tokens.length > 0) {
            runs.push(...this.processInlineTokensWithStyle((token as any).tokens, { ...baseStyle }));
          } else {
            runs.push(...this.processTextWithInlineHtml(token.text, baseStyle));
          }
          break;
        case 'strong':
          if (token.tokens && token.tokens.length > 0) {
            runs.push(...this.processInlineTokensWithStyle(token.tokens, { ...baseStyle, bold: true }));
          } else {
            runs.push(...this.processTextWithInlineHtml(token.text, { ...baseStyle, bold: true }));
          }
          break;
        case 'em':
          if (token.tokens && token.tokens.length > 0) {
            runs.push(...this.processInlineTokensWithStyle(token.tokens, { ...baseStyle, italics: true }));
          } else {
            runs.push(...this.processTextWithInlineHtml(token.text, { ...baseStyle, italics: true }));
          }
          break;
        case 'codespan':
        case 'code':
          runs.push(new TextRun({ text: this.decodeHtmlEntities(token.text), ...defaultStyle, ...baseStyle, font: 'Courier New' }));
          break;
        case 'link': {
          const rawHref = typeof token.href === 'string' ? token.href.trim() : '';
          let finalHref = rawHref;
          let finalText = this.decodeHtmlEntities(token.text || '');
          const nested = rawHref.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
          if (nested) {
            finalText = this.decodeHtmlEntities(nested[1]);
            finalHref = nested[2].trim();
          }
          const isExternal = /^(https?:\/\/|mailto:|ftp:\/\/)/i.test(finalHref);
          if (isExternal) {
            runs.push(new ExternalHyperlink({
              link: finalHref,
              children: [new TextRun({ text: finalText, ...defaultStyle, ...baseStyle, style: 'Hyperlink' })],
            }));
          } else {
            const anchorText = finalHref.replace(/^#/, '');
            const anchor = this.sanitizeBookmarkName(anchorText);
            runs.push(new InternalHyperlink({
              anchor,
              children: [new TextRun({ text: finalText, ...defaultStyle, ...baseStyle, color: '0000FF', underline: {} })],
            }));
          }
          break;
        }
        default:
          if (token.text) {
            runs.push(...this.processTextWithInlineHtml(token.text, baseStyle));
          }
      }
    });

    return runs;
  }

  // Parse inline HTML like <u>, <b>/<strong>, <i>/<em>, <s>/<del> and map to TextRun styles
  private processTextWithInlineHtml(text: string, baseStyle: any = {}): TextRun[] {
    const runs: TextRun[] = [];
    if (!text || typeof text !== 'string') {
      const defaultStyle = this.getDefaultTextStyle();
      return [new TextRun({ text: '', ...defaultStyle, ...baseStyle })];
    }

    let remaining = this.decodeHtmlEntities(text);
    const defaultStyle = this.getDefaultTextStyle();
    
    while (remaining.length > 0) {
      // Line break: <br>, <br/>, <br />
      const br = remaining.match(/^(.*?)<br\s*\/?>\s*(.*)$/i);
      if (br) {
        if (br[1]) runs.push(new TextRun({ text: br[1], ...defaultStyle, ...baseStyle }));
        runs.push(new TextRun({ break: 1 } as any));
        remaining = br[2];
        continue;
      }

      // Underline
      const u = remaining.match(/^(.*?)<u>(.*?)<\/u>(.*)$/s);
      if (u) {
        if (u[1]) runs.push(new TextRun({ text: u[1], ...defaultStyle, ...baseStyle }));
        if (u[2]) runs.push(new TextRun({ text: u[2], ...defaultStyle, ...baseStyle, underline: {} }));
        remaining = u[3];
        continue;
      }

      // Bold
      const b = remaining.match(/^(.*?)<(?:b|strong)>(.*?)<\/(?:b|strong)>(.*)$/s);
      if (b) {
        if (b[1]) runs.push(new TextRun({ text: b[1], ...defaultStyle, ...baseStyle }));
        if (b[2]) runs.push(new TextRun({ text: b[2], ...defaultStyle, ...baseStyle, bold: true }));
        remaining = b[3];
        continue;
      }

      // Italic
      const i = remaining.match(/^(.*?)<(?:i|em)>(.*?)<\/(?:i|em)>(.*)$/s);
      if (i) {
        if (i[1]) runs.push(new TextRun({ text: i[1], ...defaultStyle, ...baseStyle }));
        if (i[2]) runs.push(new TextRun({ text: i[2], ...defaultStyle, ...baseStyle, italics: true }));
        remaining = i[3];
        continue;
      }

      // Strikethrough
      const s = remaining.match(/^(.*?)<(?:s|del)>(.*?)<\/(?:s|del)>(.*)$/s);
      if (s) {
        if (s[1]) runs.push(new TextRun({ text: s[1], ...defaultStyle, ...baseStyle }));
        if (s[2]) runs.push(new TextRun({ text: s[2], ...defaultStyle, ...baseStyle, strike: true }));
        remaining = s[3];
        continue;
      }

      // No more tags
      runs.push(new TextRun({ text: remaining, ...defaultStyle, ...baseStyle }));
      break;
    }

    return runs;
  }

  private parseBasicInlineFormatting(text: string, baseStyle: any = {}): TextRun[] {
    // Approche simple et robuste pour le formatage de base
    const defaultStyle = this.getDefaultTextStyle();
    
    if (!text) {
      return [new TextRun({ text: '', ...defaultStyle, ...baseStyle })];
    }
    
    // Juste traiter **gras** de base pour commencer
    if (text.includes('**')) {
      const parts = text.split(/(\*\*[^*]+\*\*)/);
      const runs: TextRun[] = [];
      
      parts.forEach(part => {
        if (part.startsWith('**') && part.endsWith('**') && part.length > 4) {
          // Texte en gras
          const boldText = part.substring(2, part.length - 2);
          runs.push(new TextRun({ text: boldText, ...defaultStyle, ...baseStyle, bold: true }));
        } else if (part) {
          // Texte normal
          runs.push(new TextRun({ text: part, ...defaultStyle, ...baseStyle }));
        }
      });
      
      return runs.length > 0 ? runs : [new TextRun({ text, ...defaultStyle, ...baseStyle })];
    }
    
    // Traiter *italique* de base
    if (text.includes('*') && !text.includes('**')) {
      const parts = text.split(/(\*[^*]+\*)/);
      const runs: TextRun[] = [];
      
      parts.forEach(part => {
        if (part.startsWith('*') && part.endsWith('*') && part.length > 2) {
          // Texte en italique
          const italicText = part.substring(1, part.length - 1);
          runs.push(new TextRun({ text: italicText, ...defaultStyle, ...baseStyle, italics: true }));
        } else if (part) {
          // Texte normal
          runs.push(new TextRun({ text: part, ...defaultStyle, ...baseStyle }));
        }
      });
      
      return runs.length > 0 ? runs : [new TextRun({ text, ...defaultStyle, ...baseStyle })];
    }
    
    // Traiter `code` de base
    if (text.includes('`')) {
      const parts = text.split(/(`[^`]+`)/);
      const runs: TextRun[] = [];
      
      parts.forEach(part => {
        if (part.startsWith('`') && part.endsWith('`') && part.length > 2) {
          // Code
          const codeText = part.substring(1, part.length - 1);
          runs.push(new TextRun({ text: codeText, ...defaultStyle, ...baseStyle, font: 'Courier New' }));
        } else if (part) {
          // Texte normal
          runs.push(new TextRun({ text: part, ...defaultStyle, ...baseStyle }));
        }
      });
      
      return runs.length > 0 ? runs : [new TextRun({ text, ...defaultStyle, ...baseStyle })];
    }
    
    // Aucun formatage détecté
    return [new TextRun({ text, ...defaultStyle, ...baseStyle })];
  }

  private createTableOfContents(): Paragraph {
    return new Paragraph({
      children: [
        new TextRun({
          text: "Table des matières",
          bold: true,
          size: 32,
        }),
      ],
      spacing: {
        after: 400,
      },
    });
  }

  private extractHeadings(tokens: any[]): Array<{ level: number; text: string; anchor: string }> {
    const headings: Array<{ level: number; text: string; anchor: string }> = [];
    
    tokens.forEach(token => {
      if (token.type === 'heading') {
        headings.push({
          level: token.depth,
          text: token.text,
          anchor: this.generateAnchor(token.text)
        });
      }
    });
    
    return headings;
  }

  private createTableOfContentsWithLinks(): Paragraph[] {
    const tocParagraphs: Paragraph[] = [];
    
    // Title
  tocParagraphs.push(new Paragraph({
      children: [
        new TextRun({
      text: "Table des matières",
          bold: true,
          size: 32,
        }),
      ],
      spacing: {
        after: 400,
      },
    }));

    // TOC entries
  this.headings.forEach(heading => {
      const indent = (heading.level - 1) * 360; // 0.25 inch per level
      
      tocParagraphs.push(new Paragraph({
        children: [
          new InternalHyperlink({
      anchor: this.sanitizeBookmarkName(heading.anchor),
            children: [
              new TextRun({
                text: heading.text,
                color: '0000FF',
                underline: {},
              }),
            ],
          }),
        ],
        indent: {
          left: indent,
        },
        spacing: {
          after: 120,
        },
      }));
    });

    return tocParagraphs;
  }

  private createDocument(elements: (Paragraph | Table)[], options: MarkdownToDocxOptions): Document {
    // Use the stored template instead of getting it again
    const template = this.template;

    // Build numbering config: include a bullet config and one ordered config per used reference
    const orderedConfigs = Array.from(this.usedOrderedListRefs).map(ref => ({
      reference: ref,
      levels: [0,1,2].map(l => ({
        level: l,
        format: LevelFormat.DECIMAL,
        text: `%${l+1}.`,
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } }
      }))
    }));

    const bulletConfig = {
      reference: 'md-bullet',
      levels: [0,1,2].map(l => ({
        level: l,
        format: LevelFormat.BULLET,
        text: l === 0 ? '•' : l === 1 ? '◦' : '▪',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } }
      }))
    };

    return new Document({
      creator: options.author || 'Markdown-DOCX Converter',
      title: options.title || 'Converted Document',
      description: options.description || options.subject || '',
      styles: template.styles,
      numbering: {
        config: [bulletConfig, ...orderedConfigs]
      },
      sections: [{
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: elements,
      }],
    });
  }

  /** Create a fresh numbering reference for a new ordered list and track it */
  private newOrderedListReference(): string {
    this.orderedListCounter += 1;
    const ref = `md-numbered-${this.orderedListCounter}`;
    this.usedOrderedListRefs.add(ref);
    return ref;
  }

  private generateAnchor(text: string): string {
    return text
      .toLowerCase()
      .replace(/[^\w\s-]/g, '')
      .replace(/\s+/g, '-')
      .trim();
  }

  /** Ensure the anchor conforms to Word bookmark naming restrictions */
  private sanitizeBookmarkName(anchor: string): string {
    return anchor
      .replace(/^#+/, '')
      .replace(/[^a-zA-Z0-9_\-]/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '')
      .substring(0, 40) || 'a';
  }

  private getDefaultTextStyle(): any {
    // Get the default text style from template if available
    if (this.template && this.template.styles && this.template.styles.default && this.template.styles.default.document) {
      const defaultDoc = this.template.styles.default.document;
      if (defaultDoc.run) {
        return {
          font: defaultDoc.run.font,
          size: defaultDoc.run.size,
          color: defaultDoc.run.color
        };
      }
    }
    
    // Default fallback
    return {
      font: 'Calibri',
      size: 22  // 11pt in half-points
    };
  }

  private extractTextFromTokens(tokens: any[]): string {
    return tokens
      .map(token => {
        if (token.type === 'text') {
          return token.text;
        } else if (token.tokens) {
          return this.extractTextFromTokens(token.tokens);
        }
        return '';
      })
      .join(' ');
  }

  private getHeadingLevel(depth: number): typeof HeadingLevel[keyof typeof HeadingLevel] {
    switch (depth) {
      case 1: return HeadingLevel.HEADING_1;
      case 2: return HeadingLevel.HEADING_2;
      case 3: return HeadingLevel.HEADING_3;
      case 4: return HeadingLevel.HEADING_4;
      case 5: return HeadingLevel.HEADING_5;
      case 6: return HeadingLevel.HEADING_6;
      default: return HeadingLevel.HEADING_1;
    }
  }

  private countWords(text: string): number {
    return text.split(/\s+/).filter(word => word.length > 0).length;
  }

  private decodeHtmlEntities(text: string): string {
    return text
      .replace(/&#x20;/g, ' ')
  // Non-breaking hyphen
  .replace(/&#8209;/g, '‑')
      .replace(/&nbsp;/g, ' ')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'");
  }

  private parseSimpleMarkdown(text: string, baseStyle: any = {}): TextRun[] {
    if (!text || typeof text !== 'string') {
      return [new TextRun({ text: '', ...baseStyle })];
    }

    // D??coder les entit??s HTML d'abord
    text = this.decodeHtmlEntities(text);

    const runs: TextRun[] = [];
    let remaining = text;
    
    while (remaining.length > 0) {
      // Rechercher **gras**
      const boldMatch = remaining.match(/^(.*?)\*\*([^*]+)\*\*(.*)/);
      if (boldMatch) {
        // Texte avant
        if (boldMatch[1]) {
          runs.push(new TextRun({ text: boldMatch[1], ...baseStyle }));
        }
        // Texte en gras
        runs.push(new TextRun({ text: boldMatch[2], ...baseStyle, bold: true }));
        remaining = boldMatch[3];
        continue;
      }

      // Rechercher *italique* (mais pas ** qui est d??j?? trait??)
      const italicMatch = remaining.match(/^(.*?)\*([^*]+)\*(.*)/);
      if (italicMatch && !remaining.startsWith('**')) {
        // Texte avant
        if (italicMatch[1]) {
          runs.push(new TextRun({ text: italicMatch[1], ...baseStyle }));
        }
        // Texte en italique
        runs.push(new TextRun({ text: italicMatch[2], ...baseStyle, italics: true }));
        remaining = italicMatch[3];
        continue;
      }

      // Rechercher `code`
      const codeMatch = remaining.match(/^(.*?)`([^`]+)`(.*)/);
      if (codeMatch) {
        // Texte avant
        if (codeMatch[1]) {
          runs.push(new TextRun({ text: codeMatch[1], ...baseStyle }));
        }
        // Code
        runs.push(new TextRun({ text: codeMatch[2], ...baseStyle, font: 'Courier New' }));
        remaining = codeMatch[3];
        continue;
      }

      // Aucun formatage trouv??, prendre tout le texte restant
      runs.push(new TextRun({ text: remaining, ...baseStyle }));
      break;
    }

    return runs.length > 0 ? runs : [new TextRun({ text, ...baseStyle })];
  }
}
