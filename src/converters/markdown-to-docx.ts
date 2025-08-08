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
  InternalHyperlink,
  BookmarkStart,
  BookmarkEnd,
  TableOfContents,
  SectionType,
  PageBreak,
  Packer,
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

export class MarkdownToDocxConverter {
  private mermaidProcessor: MermaidPNGProcessor;
  private linkProcessor: LinkProcessor;
  private logger: Logger;
  private bookmarkCounter = 0;
  private headings: Array<{ level: number; text: string; anchor: string }> = [];
  private mermaidImages: Map<string, Buffer> = new Map();

  constructor() {
    this.mermaidProcessor = new MermaidPNGProcessor();
    this.linkProcessor = new LinkProcessor();
    this.logger = new Logger({ level: 'info' });
  }

  async convert(
    markdownContent: string,
    options: MarkdownToDocxOptions = {}
  ): Promise<ConversionResult> {
    try {
      this.logger.info('Starting Markdown to DOCX conversion');
      const startTime = Date.now();

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

      // Parse markdown
  const tokens = marked.lexer(mermaidResult.content);
      
      // Process links
      const linkResult = options.preserveLinks 
        ? this.linkProcessor.processContent(mermaidResult.content)
        : { content: mermaidResult.content, links: [], headings: [] };

      // Extract headings for TOC
      this.headings = this.extractHeadings(tokens);

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

  private async processToken(token: any): Promise<Paragraph | Table | Paragraph[] | null> {
    switch (token.type) {
      case 'heading':
        return this.createHeading(token);
      case 'paragraph':
        return this.createParagraph(token);
      case 'list':
        return this.createList(token);
      case 'blockquote':
        return this.createBlockquote(token);
      case 'code':
        return this.createCodeBlock(token); // Returns Paragraph[] now
      case 'table':
        return this.createTable(token);
      case 'hr':
        return this.createHorizontalRule();
      case 'html':
        return this.createHtml(token);
      case 'image':
        return this.createImageParagraph(token);
      case 'space':
        return null;
      default:
        this.logger.warn(`Unsupported token type: ${token.type}`);
        return null;
    }
  }

  private createHeading(token: any): Paragraph {
  const anchor = this.generateAnchor(token.text);
  this.bookmarkCounter += 1;
    const headingText = this.decodeHtmlEntities(token.text);

    return new Paragraph({
      heading: this.getHeadingLevel(token.depth),
      children: [
    // Use anchor as bookmark name so InternalHyperlink can target it
    new BookmarkStart(anchor, this.bookmarkCounter),
        new TextRun({
          text: headingText,
          bold: true,
        }),
        new BookmarkEnd(this.bookmarkCounter),
      ],
    });
  }

  private createParagraph(token: any): Paragraph {
    // Check if this paragraph contains only an image
    if (token.tokens && token.tokens.length === 1 && token.tokens[0].type === 'image') {
      return this.createImageParagraph(token.tokens[0]);
    }
    
    return new Paragraph({
      children: this.processInlineTokens(token.tokens || [{ type: 'text', text: token.text }]),
    });
  }

  private createImageParagraph(token: any): Paragraph {
    // Check if it's a Mermaid diagram reference
    if (token.href && token.href.includes('mermaid-') && token.href.endsWith('.png')) {
      // Extract the Mermaid ID from the filename
      const mermaidId = token.href.replace('mermaid-', '').replace('.png', '');
      const imageBuffer = this.mermaidImages.get(mermaidId);
      
      if (imageBuffer) {
        return new Paragraph({
          children: [
            new ImageRun({
              data: imageBuffer,
              transformation: {
                width: 600,
                height: 450,
              },
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: {
            before: 200,
            after: 200,
          },
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
          size: 24
        })
      ],
      alignment: AlignmentType.CENTER,
      spacing: {
        before: 200,
        after: 200,
      },
    });
  }

  private createList(token: any, indentLeft: number = 720): Paragraph[] {
    const paragraphs: Paragraph[] = [];

    token.items.forEach((item: any, index: number) => {
      // Puce visuelle simple ou numérotation
      const bullet = token.ordered ? `${index + 1}.` : '•';

      // Les items de liste de marked ont généralement des tokens inline (texte, strong, em, link, code...)
      const inlineTokens = Array.isArray(item.tokens)
        ? item.tokens.filter((t: any) => t.type !== 'list')
        : [{ type: 'text', text: typeof item.text === 'string' ? item.text : String(item.text || '') }];

      const formattedRuns = this.processInlineTokens(inlineTokens);

      const listChildren: TextRun[] = [
        new TextRun({ text: `${bullet} ` }),
        ...formattedRuns,
      ];

      paragraphs.push(new Paragraph({
        children: listChildren,
        indent: {
          left: indentLeft, // indentation for this list level
        },
      }));

      // Gérer d'éventuelles sous-listes imbriquées
      if (Array.isArray(item.tokens)) {
        item.tokens
          .filter((t: any) => t.type === 'list')
          .forEach((subList: any) => {
            const subParas = this.createList(subList, indentLeft + 720);
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
      const baseStyle = isHeader ? { bold: true } : {};
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

    const headerCells = token.header.map((cell: any) => new TableCell({
      children: [new Paragraph({ children: cellRuns(cell, true) })],
      shading: { fill: 'E6E6E6' },
    }));

    const rows: TableRow[] = [new TableRow({ children: headerCells })];

    token.rows.forEach((rowData: any[]) => {
      const dataCells = rowData.map((cell: any) => new TableCell({
        children: [new Paragraph({ children: cellRuns(cell, false) })],
      }));
      rows.push(new TableRow({ children: dataCells }));
    });

    return new Table({
      rows,
      width: { size: 100, type: WidthType.PERCENTAGE },
    });
  }

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
    return new Paragraph({
      children: [new TextRun({ text: '' })],
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

  private processInlineTokens(tokens: any[]): TextRun[] {
    const runs: TextRun[] = [];

    tokens.forEach(token => {
      switch (token.type) {
        case 'paragraph':
          // Some parsers (marked) wrap inline content in a paragraph token inside lists
          if (Array.isArray(token.tokens)) {
            runs.push(...this.processInlineTokens(token.tokens));
          } else if (token.text) {
            runs.push(new TextRun({ text: this.decodeHtmlEntities(token.text) }));
          }
          break;
        case 'text':
          // In marked, a text token can itself contain inline tokens (strong/em/link/codespan)
          if (Array.isArray((token as any).tokens) && (token as any).tokens.length > 0) {
            runs.push(...this.processInlineTokens((token as any).tokens));
          } else {
            runs.push(new TextRun({ text: this.decodeHtmlEntities(token.text) }));
          }
          break;
        case 'strong':
          // Pour strong, traiter les tokens enfants avec bold=true
          if (token.tokens && token.tokens.length > 0) {
            const childRuns = this.processInlineTokensWithStyle(token.tokens, { bold: true });
            runs.push(...childRuns);
          } else {
            runs.push(new TextRun({ 
              text: this.decodeHtmlEntities(token.text),
              bold: true 
            }));
          }
          break;
        case 'em':
          // Pour em, traiter les tokens enfants avec italics=true
          if (token.tokens && token.tokens.length > 0) {
            const childRuns = this.processInlineTokensWithStyle(token.tokens, { italics: true });
            runs.push(...childRuns);
          } else {
            runs.push(new TextRun({ 
              text: this.decodeHtmlEntities(token.text),
              italics: true 
            }));
          }
          break;
  case 'codespan':
  case 'code':
          runs.push(new TextRun({ 
            text: this.decodeHtmlEntities(token.text),
            font: 'Courier New'
          }));
          break;
        case 'link':
          if (token.href.startsWith('http')) {
            runs.push(new TextRun({
              text: this.decodeHtmlEntities(token.text),
              color: '0000FF',
              underline: {},
            }));
          } else {
            runs.push(new TextRun({
              text: this.decodeHtmlEntities(token.text),
              color: '0000FF',
              underline: {},
            }));
          }
          break;
        case 'image':
          // Handle image tokens - convert to placeholder text for now
          if (token.href && token.href.startsWith('data:image/')) {
            runs.push(new TextRun({
              text: `[${token.alt || 'Image'}]`,
              italics: true,
              color: '666666'
            }));
          } else {
            runs.push(new TextRun({
              text: `[Image: ${token.alt || token.href}]`,
              italics: true,
              color: '666666'
            }));
          }
          break;
        default:
          if (token.text) {
            runs.push(new TextRun({ text: token.text }));
          }
      }
    });

    return runs;
  }

  private processInlineTokensWithStyle(tokens: any[], baseStyle: any): TextRun[] {
    const runs: TextRun[] = [];

    tokens.forEach(token => {
      switch (token.type) {
        case 'paragraph':
          if (Array.isArray(token.tokens)) {
            runs.push(...this.processInlineTokensWithStyle(token.tokens, { ...baseStyle }));
          } else if (token.text) {
            runs.push(new TextRun({ text: token.text, ...baseStyle }));
          }
          break;
        case 'text':
          if (Array.isArray((token as any).tokens) && (token as any).tokens.length > 0) {
            runs.push(...this.processInlineTokensWithStyle((token as any).tokens, { ...baseStyle }));
          } else {
            runs.push(new TextRun({ 
              text: token.text,
              ...baseStyle
            }));
          }
          break;
        case 'strong':
          // Combiner bold avec le style de base
          if (token.tokens && token.tokens.length > 0) {
            const childRuns = this.processInlineTokensWithStyle(token.tokens, { ...baseStyle, bold: true });
            runs.push(...childRuns);
          } else {
            runs.push(new TextRun({ 
              text: token.text,
              ...baseStyle,
              bold: true 
            }));
          }
          break;
        case 'em':
          // Combiner italics avec le style de base
          if (token.tokens && token.tokens.length > 0) {
            const childRuns = this.processInlineTokensWithStyle(token.tokens, { ...baseStyle, italics: true });
            runs.push(...childRuns);
          } else {
            runs.push(new TextRun({ 
              text: token.text,
              ...baseStyle,
              italics: true 
            }));
          }
          break;
  case 'codespan':
  case 'code':
          runs.push(new TextRun({ 
            text: token.text,
            ...baseStyle,
            font: 'Courier New'
          }));
          break;
        default:
          if (token.text) {
            runs.push(new TextRun({ 
              text: token.text,
              ...baseStyle
            }));
          }
      }
    });

    return runs;
  }

  private parseBasicInlineFormatting(text: string, baseStyle: any = {}): TextRun[] {
    // Approche simple et robuste pour le formatage de base
    if (!text) {
      return [new TextRun({ text: '', ...baseStyle })];
    }
    
    // Juste traiter **gras** de base pour commencer
    if (text.includes('**')) {
      const parts = text.split(/(\*\*[^*]+\*\*)/);
      const runs: TextRun[] = [];
      
      parts.forEach(part => {
        if (part.startsWith('**') && part.endsWith('**') && part.length > 4) {
          // Texte en gras
          const boldText = part.substring(2, part.length - 2);
          runs.push(new TextRun({ text: boldText, ...baseStyle, bold: true }));
        } else if (part) {
          // Texte normal
          runs.push(new TextRun({ text: part, ...baseStyle }));
        }
      });
      
      return runs.length > 0 ? runs : [new TextRun({ text, ...baseStyle })];
    }
    
    // Traiter *italique* de base
    if (text.includes('*') && !text.includes('**')) {
      const parts = text.split(/(\*[^*]+\*)/);
      const runs: TextRun[] = [];
      
      parts.forEach(part => {
        if (part.startsWith('*') && part.endsWith('*') && part.length > 2) {
          // Texte en italique
          const italicText = part.substring(1, part.length - 1);
          runs.push(new TextRun({ text: italicText, ...baseStyle, italics: true }));
        } else if (part) {
          // Texte normal
          runs.push(new TextRun({ text: part, ...baseStyle }));
        }
      });
      
      return runs.length > 0 ? runs : [new TextRun({ text, ...baseStyle })];
    }
    
    // Traiter `code` de base
    if (text.includes('`')) {
      const parts = text.split(/(`[^`]+`)/);
      const runs: TextRun[] = [];
      
      parts.forEach(part => {
        if (part.startsWith('`') && part.endsWith('`') && part.length > 2) {
          // Code
          const codeText = part.substring(1, part.length - 1);
          runs.push(new TextRun({ text: codeText, ...baseStyle, font: 'Courier New' }));
        } else if (part) {
          // Texte normal
          runs.push(new TextRun({ text: part, ...baseStyle }));
        }
      });
      
      return runs.length > 0 ? runs : [new TextRun({ text, ...baseStyle })];
    }
    
    // Aucun formatage d??tect??
    return [new TextRun({ text, ...baseStyle })];
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
            anchor: heading.anchor,
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
    const template = DocumentTemplates.getTemplate(options.template || 'professional-report');
    
    return new Document({
      creator: options.author || 'Markdown-DOCX Converter',
      title: options.title || 'Converted Document',
      description: options.subject || '',
      styles: template.styles,
      sections: [{
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: elements,
      }],
    });
  }

  private generateAnchor(text: string): string {
    return text
      .toLowerCase()
      .replace(/[^\w\s-]/g, '')
      .replace(/\s+/g, '-')
      .trim();
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
