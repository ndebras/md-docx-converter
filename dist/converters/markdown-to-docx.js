"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MarkdownToDocxConverter = void 0;
const docx_1 = require("docx");
const marked_1 = require("marked");
const mermaid_png_processor_1 = require("../processors/mermaid-png-processor");
const link_processor_1 = require("../processors/link-processor");
const document_templates_1 = require("../styles/document-templates");
const logger_1 = require("../utils/logger");
class MarkdownToDocxConverter {
    constructor() {
        this.bookmarkCounter = 0;
        this.headings = [];
        this.mermaidImages = new Map();
        this.mermaidProcessor = new mermaid_png_processor_1.MermaidPNGProcessor();
        this.linkProcessor = new link_processor_1.LinkProcessor();
        this.logger = new logger_1.Logger({ level: 'info' });
    }
    async convert(markdownContent, options = {}) {
        try {
            this.logger.info('Starting Markdown to DOCX conversion');
            const startTime = Date.now();
            // Process Mermaid diagrams first and get PNG images
            const mermaidResult = await this.mermaidProcessor.processContent(markdownContent);
            // Store images for later use in DOCX
            mermaidResult.images.forEach(image => {
                this.mermaidImages.set(image.id, image.buffer);
            });
            // Parse markdown
            const tokens = marked_1.marked.lexer(mermaidResult.content);
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
            const document = this.createDocument(elements, options);
            // Generate buffer
            const buffer = await docx_1.Packer.toBuffer(document);
            const processingTime = Date.now() - startTime;
            const metadata = {
                inputSize: markdownContent.length,
                outputSize: buffer.length,
                processingTime,
                pageCount: Math.ceil(buffer.length / 4000), // Rough estimation
                mermaidDiagramCount: mermaidResult.diagramCount || 0,
                internalLinkCount: linkResult.links.filter((l) => l.type === 'internal').length,
                externalLinkCount: linkResult.links.filter((l) => l.type === 'external').length,
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
        }
        catch (error) {
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
    async processTokens(tokens) {
        const elements = [];
        for (const token of tokens) {
            const result = await this.processToken(token);
            if (result) {
                if (Array.isArray(result)) {
                    elements.push(...result);
                }
                else {
                    elements.push(result);
                }
            }
        }
        return elements;
    }
    async processToken(token) {
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
    createHeading(token) {
        const bookmarkId = `heading_${++this.bookmarkCounter}`;
        const anchor = this.generateAnchor(token.text);
        const headingText = this.decodeHtmlEntities(token.text);
        return new docx_1.Paragraph({
            heading: this.getHeadingLevel(token.depth),
            children: [
                new docx_1.BookmarkStart(bookmarkId, this.bookmarkCounter),
                new docx_1.TextRun({
                    text: headingText,
                    bold: true,
                }),
                new docx_1.BookmarkEnd(this.bookmarkCounter),
            ],
        });
    }
    createParagraph(token) {
        // Check if this paragraph contains only an image
        if (token.tokens && token.tokens.length === 1 && token.tokens[0].type === 'image') {
            return this.createImageParagraph(token.tokens[0]);
        }
        return new docx_1.Paragraph({
            children: this.processInlineTokens(token.tokens || [{ type: 'text', text: token.text }]),
        });
    }
    createImageParagraph(token) {
        // Check if it's a Mermaid diagram reference
        if (token.href && token.href.includes('mermaid-') && token.href.endsWith('.png')) {
            // Extract the Mermaid ID from the filename
            const mermaidId = token.href.replace('mermaid-', '').replace('.png', '');
            const imageBuffer = this.mermaidImages.get(mermaidId);
            if (imageBuffer) {
                return new docx_1.Paragraph({
                    children: [
                        new docx_1.ImageRun({
                            data: imageBuffer,
                            transformation: {
                                width: 600,
                                height: 450,
                            },
                        }),
                    ],
                    alignment: docx_1.AlignmentType.CENTER,
                    spacing: {
                        before: 200,
                        after: 200,
                    },
                });
            }
        }
        // Fallback for other images or missing Mermaid images
        return new docx_1.Paragraph({
            children: [
                new docx_1.TextRun({
                    text: `[Image: ${token.alt || token.href || 'Non disponible'}]`,
                    italics: true,
                    color: '0066CC',
                    size: 24
                })
            ],
            alignment: docx_1.AlignmentType.CENTER,
            spacing: {
                before: 200,
                after: 200,
            },
        });
    }
    createList(token) {
        const paragraphs = [];
        token.items.forEach((item, index) => {
            const bullet = token.ordered ? `${index + 1}.` : '???';
            let itemText = '';
            // Extraire le texte de l'item en g??rant les diff??rents formats
            if (typeof item.text === 'string') {
                itemText = item.text;
            }
            else if (item.tokens && Array.isArray(item.tokens)) {
                // Extraire le texte des tokens
                itemText = this.extractTextFromTokens(item.tokens);
            }
            else {
                itemText = String(item.text || '');
            }
            // Traiter le formatage Markdown dans le texte de l'item
            const formattedRuns = this.parseSimpleMarkdown(itemText);
            // Cr??er les enfants avec la puce + le texte format??
            const listChildren = [
                new docx_1.TextRun({ text: `${bullet} ` }),
                ...formattedRuns
            ];
            paragraphs.push(new docx_1.Paragraph({
                children: listChildren,
                indent: {
                    left: 720, // 0.5 inch
                },
            }));
        });
        return paragraphs;
    }
    createBlockquote(token) {
        const paragraphs = [];
        if (token.tokens) {
            token.tokens.forEach((subToken) => {
                if (subToken.type === 'paragraph') {
                    paragraphs.push(new docx_1.Paragraph({
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
    createCodeBlock(token) {
        // Split the code into lines and create a separate paragraph for each line
        const lines = token.text.split('\n');
        const paragraphs = [];
        lines.forEach((line, index) => {
            paragraphs.push(new docx_1.Paragraph({
                children: [
                    new docx_1.TextRun({
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
    createTable(token) {
        const headerCells = token.header.map((cell) => {
            const cellText = String(cell);
            const formattedRuns = this.parseSimpleMarkdown(cellText, { bold: true });
            return new docx_1.TableCell({
                children: [new docx_1.Paragraph({
                        children: formattedRuns
                    })],
                shading: { fill: 'E6E6E6' },
            });
        });
        const rows = [new docx_1.TableRow({ children: headerCells })];
        token.rows.forEach((rowData) => {
            const dataCells = rowData.map((cell) => {
                const cellText = String(cell);
                const formattedRuns = this.parseSimpleMarkdown(cellText);
                return new docx_1.TableCell({
                    children: [new docx_1.Paragraph({
                            children: formattedRuns
                        })],
                });
            });
            rows.push(new docx_1.TableRow({ children: dataCells }));
        });
        return new docx_1.Table({
            rows,
            width: {
                size: 100,
                type: docx_1.WidthType.PERCENTAGE,
            },
        });
    }
    createHtml(token) {
        // Handle Mermaid diagrams
        if (token.text.includes('mermaid')) {
            // This should be handled by the MermaidProcessor
            return null;
        }
        // Basic HTML handling
        const text = token.text.replace(/<[^>]*>/g, '');
        return new docx_1.Paragraph({
            children: [new docx_1.TextRun({ text })],
        });
    }
    createHorizontalRule() {
        return new docx_1.Paragraph({
            children: [new docx_1.TextRun({ text: '' })],
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
    processInlineTokens(tokens) {
        const runs = [];
        tokens.forEach(token => {
            switch (token.type) {
                case 'text':
                    runs.push(new docx_1.TextRun({ text: this.decodeHtmlEntities(token.text) }));
                    break;
                case 'strong':
                    // Pour strong, traiter les tokens enfants avec bold=true
                    if (token.tokens && token.tokens.length > 0) {
                        const childRuns = this.processInlineTokensWithStyle(token.tokens, { bold: true });
                        runs.push(...childRuns);
                    }
                    else {
                        runs.push(new docx_1.TextRun({
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
                    }
                    else {
                        runs.push(new docx_1.TextRun({
                            text: this.decodeHtmlEntities(token.text),
                            italics: true
                        }));
                    }
                    break;
                case 'code':
                    runs.push(new docx_1.TextRun({
                        text: this.decodeHtmlEntities(token.text),
                        font: 'Courier New'
                    }));
                    break;
                case 'link':
                    if (token.href.startsWith('http')) {
                        runs.push(new docx_1.TextRun({
                            text: this.decodeHtmlEntities(token.text),
                            color: '0000FF',
                            underline: {},
                        }));
                    }
                    else {
                        runs.push(new docx_1.TextRun({
                            text: this.decodeHtmlEntities(token.text),
                            color: '0000FF',
                            underline: {},
                        }));
                    }
                    break;
                case 'image':
                    // Handle image tokens - convert to placeholder text for now
                    if (token.href && token.href.startsWith('data:image/')) {
                        runs.push(new docx_1.TextRun({
                            text: `[${token.alt || 'Image'}]`,
                            italics: true,
                            color: '666666'
                        }));
                    }
                    else {
                        runs.push(new docx_1.TextRun({
                            text: `[Image: ${token.alt || token.href}]`,
                            italics: true,
                            color: '666666'
                        }));
                    }
                    break;
                default:
                    if (token.text) {
                        runs.push(new docx_1.TextRun({ text: token.text }));
                    }
            }
        });
        return runs;
    }
    processInlineTokensWithStyle(tokens, baseStyle) {
        const runs = [];
        tokens.forEach(token => {
            switch (token.type) {
                case 'text':
                    runs.push(new docx_1.TextRun({
                        text: token.text,
                        ...baseStyle
                    }));
                    break;
                case 'strong':
                    // Combiner bold avec le style de base
                    if (token.tokens && token.tokens.length > 0) {
                        const childRuns = this.processInlineTokensWithStyle(token.tokens, { ...baseStyle, bold: true });
                        runs.push(...childRuns);
                    }
                    else {
                        runs.push(new docx_1.TextRun({
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
                    }
                    else {
                        runs.push(new docx_1.TextRun({
                            text: token.text,
                            ...baseStyle,
                            italics: true
                        }));
                    }
                    break;
                case 'code':
                    runs.push(new docx_1.TextRun({
                        text: token.text,
                        ...baseStyle,
                        font: 'Courier New'
                    }));
                    break;
                default:
                    if (token.text) {
                        runs.push(new docx_1.TextRun({
                            text: token.text,
                            ...baseStyle
                        }));
                    }
            }
        });
        return runs;
    }
    parseBasicInlineFormatting(text, baseStyle = {}) {
        // Approche simple et robuste pour le formatage de base
        if (!text) {
            return [new docx_1.TextRun({ text: '', ...baseStyle })];
        }
        // Juste traiter **gras** de base pour commencer
        if (text.includes('**')) {
            const parts = text.split(/(\*\*[^*]+\*\*)/);
            const runs = [];
            parts.forEach(part => {
                if (part.startsWith('**') && part.endsWith('**') && part.length > 4) {
                    // Texte en gras
                    const boldText = part.substring(2, part.length - 2);
                    runs.push(new docx_1.TextRun({ text: boldText, ...baseStyle, bold: true }));
                }
                else if (part) {
                    // Texte normal
                    runs.push(new docx_1.TextRun({ text: part, ...baseStyle }));
                }
            });
            return runs.length > 0 ? runs : [new docx_1.TextRun({ text, ...baseStyle })];
        }
        // Traiter *italique* de base
        if (text.includes('*') && !text.includes('**')) {
            const parts = text.split(/(\*[^*]+\*)/);
            const runs = [];
            parts.forEach(part => {
                if (part.startsWith('*') && part.endsWith('*') && part.length > 2) {
                    // Texte en italique
                    const italicText = part.substring(1, part.length - 1);
                    runs.push(new docx_1.TextRun({ text: italicText, ...baseStyle, italics: true }));
                }
                else if (part) {
                    // Texte normal
                    runs.push(new docx_1.TextRun({ text: part, ...baseStyle }));
                }
            });
            return runs.length > 0 ? runs : [new docx_1.TextRun({ text, ...baseStyle })];
        }
        // Traiter `code` de base
        if (text.includes('`')) {
            const parts = text.split(/(`[^`]+`)/);
            const runs = [];
            parts.forEach(part => {
                if (part.startsWith('`') && part.endsWith('`') && part.length > 2) {
                    // Code
                    const codeText = part.substring(1, part.length - 1);
                    runs.push(new docx_1.TextRun({ text: codeText, ...baseStyle, font: 'Courier New' }));
                }
                else if (part) {
                    // Texte normal
                    runs.push(new docx_1.TextRun({ text: part, ...baseStyle }));
                }
            });
            return runs.length > 0 ? runs : [new docx_1.TextRun({ text, ...baseStyle })];
        }
        // Aucun formatage d??tect??
        return [new docx_1.TextRun({ text, ...baseStyle })];
    }
    createTableOfContents() {
        return new docx_1.Paragraph({
            children: [
                new docx_1.TextRun({
                    text: "Table des mati??res",
                    bold: true,
                    size: 32,
                }),
            ],
            spacing: {
                after: 400,
            },
        });
    }
    extractHeadings(tokens) {
        const headings = [];
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
    createTableOfContentsWithLinks() {
        const tocParagraphs = [];
        // Title
        tocParagraphs.push(new docx_1.Paragraph({
            children: [
                new docx_1.TextRun({
                    text: "Table des mati??res",
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
            tocParagraphs.push(new docx_1.Paragraph({
                children: [
                    new docx_1.InternalHyperlink({
                        anchor: heading.anchor,
                        children: [
                            new docx_1.TextRun({
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
    createDocument(elements, options) {
        const template = document_templates_1.DocumentTemplates.getTemplate(options.template || 'professional-report');
        return new docx_1.Document({
            creator: options.author || 'Markdown-DOCX Converter',
            title: options.title || 'Converted Document',
            description: options.subject || '',
            styles: template.styles,
            sections: [{
                    properties: {
                        type: docx_1.SectionType.CONTINUOUS,
                    },
                    children: elements,
                }],
        });
    }
    generateAnchor(text) {
        return text
            .toLowerCase()
            .replace(/[^\w\s-]/g, '')
            .replace(/\s+/g, '-')
            .trim();
    }
    extractTextFromTokens(tokens) {
        return tokens
            .map(token => {
            if (token.type === 'text') {
                return token.text;
            }
            else if (token.tokens) {
                return this.extractTextFromTokens(token.tokens);
            }
            return '';
        })
            .join(' ');
    }
    getHeadingLevel(depth) {
        switch (depth) {
            case 1: return docx_1.HeadingLevel.HEADING_1;
            case 2: return docx_1.HeadingLevel.HEADING_2;
            case 3: return docx_1.HeadingLevel.HEADING_3;
            case 4: return docx_1.HeadingLevel.HEADING_4;
            case 5: return docx_1.HeadingLevel.HEADING_5;
            case 6: return docx_1.HeadingLevel.HEADING_6;
            default: return docx_1.HeadingLevel.HEADING_1;
        }
    }
    countWords(text) {
        return text.split(/\s+/).filter(word => word.length > 0).length;
    }
    decodeHtmlEntities(text) {
        return text
            .replace(/&#x20;/g, ' ')
            .replace(/&#8209;/g, '???')
            .replace(/&nbsp;/g, ' ')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .replace(/&quot;/g, '"')
            .replace(/&#39;/g, "'");
    }
    parseSimpleMarkdown(text, baseStyle = {}) {
        if (!text || typeof text !== 'string') {
            return [new docx_1.TextRun({ text: '', ...baseStyle })];
        }
        // D??coder les entit??s HTML d'abord
        text = this.decodeHtmlEntities(text);
        const runs = [];
        let remaining = text;
        while (remaining.length > 0) {
            // Rechercher **gras**
            const boldMatch = remaining.match(/^(.*?)\*\*([^*]+)\*\*(.*)/);
            if (boldMatch) {
                // Texte avant
                if (boldMatch[1]) {
                    runs.push(new docx_1.TextRun({ text: boldMatch[1], ...baseStyle }));
                }
                // Texte en gras
                runs.push(new docx_1.TextRun({ text: boldMatch[2], ...baseStyle, bold: true }));
                remaining = boldMatch[3];
                continue;
            }
            // Rechercher *italique* (mais pas ** qui est d??j?? trait??)
            const italicMatch = remaining.match(/^(.*?)\*([^*]+)\*(.*)/);
            if (italicMatch && !remaining.startsWith('**')) {
                // Texte avant
                if (italicMatch[1]) {
                    runs.push(new docx_1.TextRun({ text: italicMatch[1], ...baseStyle }));
                }
                // Texte en italique
                runs.push(new docx_1.TextRun({ text: italicMatch[2], ...baseStyle, italics: true }));
                remaining = italicMatch[3];
                continue;
            }
            // Rechercher `code`
            const codeMatch = remaining.match(/^(.*?)`([^`]+)`(.*)/);
            if (codeMatch) {
                // Texte avant
                if (codeMatch[1]) {
                    runs.push(new docx_1.TextRun({ text: codeMatch[1], ...baseStyle }));
                }
                // Code
                runs.push(new docx_1.TextRun({ text: codeMatch[2], ...baseStyle, font: 'Courier New' }));
                remaining = codeMatch[3];
                continue;
            }
            // Aucun formatage trouv??, prendre tout le texte restant
            runs.push(new docx_1.TextRun({ text: remaining, ...baseStyle }));
            break;
        }
        return runs.length > 0 ? runs : [new docx_1.TextRun({ text, ...baseStyle })];
    }
}
exports.MarkdownToDocxConverter = MarkdownToDocxConverter;
//# sourceMappingURL=markdown-to-docx.js.map