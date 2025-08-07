"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MarkdownToDocxConverter = void 0;
const docx_1 = require("docx");
const marked_1 = require("marked");
const document_templates_1 = require("../styles/document-templates");
const mermaid_processor_1 = require("../processors/mermaid-processor");
const link_processor_1 = require("../processors/link-processor");
const utils_1 = require("../utils");
const logger_1 = require("../utils/logger");
/**
 * Advanced Markdown to DOCX converter with professional styling
 */
class MarkdownToDocxConverter {
    constructor(options = {}) {
        this.processedDiagrams = [];
        this.bookmarkCounter = 0;
        this.mermaidProcessor = new mermaid_processor_1.MermaidProcessor(options.mermaidTheme);
        this.linkProcessor = new link_processor_1.LinkProcessor();
        this.templateConfig = document_templates_1.DocumentTemplates.getTemplate(options.template || 'professional-report');
    }
    /**
     * Convert Markdown to DOCX
     */
    async convert(markdownContent, options = {}) {
        const startTime = Date.now();
        try {
            logger_1.logger.info('Starting Markdown to DOCX conversion');
            // Validate input
            if (!markdownContent.trim()) {
                throw new Error('Empty markdown content provided');
            }
            // Initialize processors
            await this.mermaidProcessor.initialize();
            // Process Mermaid diagrams
            const { result: processedMarkdown, elapsed: mermaidTime } = await utils_1.PerformanceUtils.measureAsync(async () => {
                return await this.processMermaidDiagrams(markdownContent);
            }, 'mermaid-processing');
            logger_1.logger.debug(`Mermaid processing completed in ${mermaidTime}ms`);
            // Process links
            const { processedContent: linkedMarkdown, links } = this.linkProcessor.processMarkdownLinks(processedMarkdown);
            // Parse markdown to tokens
            const tokens = marked_1.marked.lexer(linkedMarkdown);
            logger_1.logger.debug(`Parsed ${tokens.length} markdown tokens`);
            // Generate table of contents if requested
            let tocElements = [];
            if (options.tocGeneration) {
                const { tocMarkdown } = this.linkProcessor.generateTableOfContents(markdownContent);
                const tocTokens = marked_1.marked.lexer(tocMarkdown);
                tocElements = await this.processTokens(tocTokens);
                tocElements.push(new docx_1.Paragraph({ children: [new docx_1.PageBreak()] }));
            }
            // Convert tokens to DOCX elements
            const { result: documentElements, elapsed: conversionTime } = await utils_1.PerformanceUtils.measureAsync(async () => {
                return await this.processTokens(tokens);
            }, 'token-conversion');
            logger_1.logger.debug(`Token conversion completed in ${conversionTime}ms`);
            // Create document
            const document = this.createDocument([...tocElements, ...documentElements], options);
            // Generate buffer
            const { result: buffer, elapsed: renderTime } = await utils_1.PerformanceUtils.measureAsync(async () => {
                return await document.build();
            }, 'document-render');
            logger_1.logger.debug(`Document rendering completed in ${renderTime}ms`);
            const totalTime = Date.now() - startTime;
            const metadata = {
                inputSize: Buffer.from(markdownContent, 'utf-8').length,
                outputSize: buffer.length,
                processingTime: totalTime,
                mermaidDiagramCount: this.processedDiagrams.length,
                internalLinkCount: links.filter(l => l.type === 'internal' || l.type === 'anchor').length,
                externalLinkCount: links.filter(l => l.type === 'external').length,
            };
            logger_1.logger.info('Markdown to DOCX conversion completed successfully', metadata);
            return {
                success: true,
                output: buffer,
                metadata,
                warnings: this.getConversionWarnings(),
            };
        }
        catch (error) {
            logger_1.logger.error('Markdown to DOCX conversion failed', { error });
            return {
                success: false,
                error: {
                    code: 'CONVERSION_FAILED',
                    message: error instanceof Error ? error.message : String(error),
                    details: { processingTime: Date.now() - startTime },
                },
            };
        }
        finally {
            await this.cleanup();
        }
    }
    /**
     * Process Mermaid diagrams
     */
    async processMermaidDiagrams(markdown) {
        const diagrams = this.mermaidProcessor.extractMermaidDiagrams(markdown);
        this.processedDiagrams = [];
        if (diagrams.length === 0) {
            return markdown;
        }
        logger_1.logger.info(`Processing ${diagrams.length} Mermaid diagrams`);
        for (const diagram of diagrams) {
            try {
                const processedDiagram = await this.mermaidProcessor.processDiagram(diagram.code, 'png', // Use PNG for better Word compatibility
                { width: 800, height: 600, backgroundColor: 'white' });
                this.processedDiagrams.push(processedDiagram);
                logger_1.logger.debug(`Processed Mermaid diagram: ${diagram.id}`);
            }
            catch (error) {
                logger_1.logger.warn(`Failed to process Mermaid diagram: ${diagram.id}`, { error });
            }
        }
        // Replace Mermaid blocks with placeholder text
        return this.mermaidProcessor.replaceMermaidBlocks(markdown, this.processedDiagrams, './diagrams');
    }
    /**
     * Process markdown tokens to DOCX elements
     */
    async processTokens(tokens) {
        const elements = [];
        for (const token of tokens) {
            try {
                const element = await this.processToken(token);
                if (element) {
                    if (Array.isArray(element)) {
                        elements.push(...element);
                    }
                    else {
                        elements.push(element);
                    }
                }
            }
            catch (error) {
                logger_1.logger.warn('Failed to process token', { token: token.type, error });
            }
        }
        return elements;
    }
    /**
     * Process individual markdown token
     */
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
                return this.createCodeBlock(token);
            case 'table':
                return this.createTable(token);
            case 'hr':
                return this.createHorizontalRule();
            case 'html':
                return this.createHtml(token);
            case 'space':
                return null; // Skip spaces
            default:
                logger_1.logger.debug(`Unhandled token type: ${token.type}`);
                return null;
        }
    }
    /**
     * Create heading element
     */
    createHeading(token) {
        const anchor = utils_1.StringUtils.createAnchor(token.text);
        const bookmarkId = `heading_${this.bookmarkCounter++}`;
        const headingStyle = this.getHeadingStyle(token.depth);
        return new docx_1.Paragraph({
            children: [
                new docx_1.BookmarkStart({
                    id: bookmarkId,
                    name: anchor,
                }),
                new docx_1.TextRun({
                    text: token.text,
                    ...headingStyle.run,
                }),
                new docx_1.BookmarkEnd({
                    id: bookmarkId,
                }),
            ],
            heading: this.getHeadingLevel(token.depth),
            ...headingStyle.paragraph,
        });
    }
    /**
     * Create paragraph element
     */
    createParagraph(token) {
        const children = this.processInlineTokens(token.tokens);
        return new docx_1.Paragraph({
            children,
            ...this.templateConfig.styles.paragraph?.normal?.paragraph,
        });
    }
    /**
     * Create list elements
     */
    createList(token) {
        const paragraphs = [];
        for (let i = 0; i < token.items.length; i++) {
            const item = token.items[i];
            const children = this.processInlineTokens(item.tokens);
            const bullet = token.ordered ? `${i + 1}.` : '???';
            paragraphs.push(new docx_1.Paragraph({
                children: [
                    new docx_1.TextRun({
                        text: `${bullet} `,
                        ...this.templateConfig.styles.default.document.run,
                    }),
                    ...children,
                ],
                indent: {
                    left: 720, // 0.5 inch
                    hanging: 360, // 0.25 inch
                },
                ...this.templateConfig.styles.paragraph?.normal?.paragraph,
            }));
        }
        return paragraphs;
    }
    /**
     * Create blockquote element
     */
    createBlockquote(token) {
        const paragraphs = [];
        for (const subToken of token.tokens) {
            if (subToken.type === 'paragraph') {
                const children = this.processInlineTokens(subToken.tokens);
                paragraphs.push(new docx_1.Paragraph({
                    children,
                    indent: {
                        left: 720, // 0.5 inch
                    },
                    border: {
                        left: {
                            color: 'CCCCCC',
                            space: 1,
                            style: 'single',
                            size: 6,
                        },
                    },
                    spacing: {
                        before: 120,
                        after: 120,
                    },
                }));
            }
        }
        return paragraphs;
    }
    /**
     * Create code block element
     */
    createCodeBlock(token) {
        return new docx_1.Paragraph({
            children: [
                new docx_1.TextRun({
                    text: token.text,
                    ...this.templateConfig.styles.codeBlock?.run,
                }),
            ],
            ...this.templateConfig.styles.codeBlock?.paragraph,
        });
    }
    /**
     * Create table element
     */
    createTable(token) {
        const rows = [];
        // Header row
        if (token.header.length > 0) {
            const headerCells = token.header.map(cell => new docx_1.TableCell({
                children: [
                    new docx_1.Paragraph({
                        children: this.processInlineTokens(cell.tokens),
                        ...this.templateConfig.styles.table?.headerStyle,
                    }),
                ],
                shading: {
                    fill: 'F0F0F0',
                },
            }));
            rows.push(new docx_1.TableRow({ children: headerCells }));
        }
        // Data rows
        for (const rowData of token.rows) {
            const dataCells = rowData.map(cell => new docx_1.TableCell({
                children: [
                    new docx_1.Paragraph({
                        children: this.processInlineTokens(cell.tokens),
                    }),
                ],
            }));
            rows.push(new docx_1.TableRow({ children: dataCells }));
        }
        const table = new docx_1.Table({
            rows,
            ...this.templateConfig.styles.table,
        });
        // Return as paragraph container
        return new docx_1.Paragraph({
            children: [],
            spacing: { before: 120, after: 120 },
        });
    }
    /**
     * Create horizontal rule
     */
    createHorizontalRule() {
        return new docx_1.Paragraph({
            children: [],
            border: {
                bottom: {
                    color: 'CCCCCC',
                    space: 1,
                    style: 'single',
                    size: 6,
                },
            },
            spacing: {
                before: 240,
                after: 240,
            },
        });
    }
    /**
     * Create HTML element (limited support)
     */
    createHtml(token) {
        // Basic HTML to text conversion
        const text = utils_1.StringUtils.stripHtml(token.text);
        if (text.trim()) {
            return new docx_1.Paragraph({
                children: [
                    new docx_1.TextRun({
                        text,
                        ...this.templateConfig.styles.default.document.run,
                    }),
                ],
            });
        }
        return null;
    }
    /**
     * Process inline tokens (bold, italic, links, etc.)
     */
    processInlineTokens(tokens) {
        const runs = [];
        for (const token of tokens) {
            switch (token.type) {
                case 'text':
                    runs.push(new docx_1.TextRun({
                        text: token.text,
                        ...this.templateConfig.styles.default.document.run,
                    }));
                    break;
                case 'strong':
                    const strongText = this.extractTextFromTokens(token.tokens);
                    runs.push(new docx_1.TextRun({
                        text: strongText,
                        bold: true,
                        ...this.templateConfig.styles.default.document.run,
                    }));
                    break;
                case 'em':
                    const emText = this.extractTextFromTokens(token.tokens);
                    runs.push(new docx_1.TextRun({
                        text: emText,
                        italics: true,
                        ...this.templateConfig.styles.default.document.run,
                    }));
                    break;
                case 'code':
                    runs.push(new docx_1.TextRun({
                        text: token.text,
                        font: 'Consolas',
                        size: 20,
                        shading: {
                            fill: 'F0F0F0',
                        },
                    }));
                    break;
                case 'link':
                    const linkText = this.extractTextFromTokens(token.tokens);
                    if (token.href.startsWith('#')) {
                        // Internal link
                        runs.push(new docx_1.TextRun({
                            text: linkText,
                            ...this.templateConfig.styles.hyperlink?.run,
                        }));
                    }
                    else {
                        // External link
                        runs.push(new docx_1.TextRun({
                            text: linkText,
                            ...this.templateConfig.styles.hyperlink?.run,
                        }));
                    }
                    break;
                case 'image':
                    // Handle images (if we have processed Mermaid diagrams)
                    const imageDiagram = this.processedDiagrams.find(d => token.href.includes(d.id));
                    if (imageDiagram) {
                        runs.push(new docx_1.TextRun({
                            text: `[${token.text || 'Diagram'}]`,
                            ...this.templateConfig.styles.default.document.run,
                        }));
                    }
                    else {
                        runs.push(new docx_1.TextRun({
                            text: `[Image: ${token.text || token.href}]`,
                            ...this.templateConfig.styles.default.document.run,
                        }));
                    }
                    break;
                default:
                    logger_1.logger.debug(`Unhandled inline token type: ${token.type}`);
                    break;
            }
        }
        return runs;
    }
    /**
     * Extract plain text from tokens
     */
    extractTextFromTokens(tokens) {
        return tokens
            .map(token => {
            if ('text' in token) {
                return token.text;
            }
            if ('tokens' in token && Array.isArray(token.tokens)) {
                return this.extractTextFromTokens(token.tokens);
            }
            return '';
        })
            .join('');
    }
    /**
     * Get heading style based on depth
     */
    getHeadingStyle(depth) {
        const headingMap = {
            1: 'heading1',
            2: 'heading2',
            3: 'heading3',
            4: 'heading3', // Use heading3 style for h4+
            5: 'heading3',
            6: 'heading3',
        };
        const styleName = headingMap[depth];
        return this.templateConfig.styles.headings[styleName] || this.templateConfig.styles.headings.heading3;
    }
    /**
     * Get DOCX heading level
     */
    getHeadingLevel(depth) {
        const levelMap = {
            1: docx_1.HeadingLevel.HEADING_1,
            2: docx_1.HeadingLevel.HEADING_2,
            3: docx_1.HeadingLevel.HEADING_3,
            4: docx_1.HeadingLevel.HEADING_4,
            5: docx_1.HeadingLevel.HEADING_5,
            6: docx_1.HeadingLevel.HEADING_6,
        };
        return levelMap[depth] || docx_1.HeadingLevel.HEADING_6;
    }
    /**
     * Create final document
     */
    createDocument(elements, options) {
        const metadata = {
            title: options.title || 'Document',
            subject: options.subject || '',
            creator: options.author || 'Markdown-DOCX Converter',
            keywords: options.keywords?.join(', ') || '',
            description: options.description || '',
            created: options.createdAt || new Date(),
            modified: new Date(),
        };
        return new docx_1.Document({
            creator: metadata.creator,
            title: metadata.title,
            subject: metadata.subject,
            keywords: metadata.keywords,
            description: metadata.description,
            sections: [
                {
                    properties: {
                        ...this.templateConfig.sections.properties,
                    },
                    headers: this.templateConfig.sections.headers || {},
                    footers: this.templateConfig.sections.footers || {},
                    children: elements,
                },
            ],
        });
    }
    /**
     * Get conversion warnings
     */
    getConversionWarnings() {
        const warnings = [];
        if (this.processedDiagrams.length > 0) {
            warnings.push(`${this.processedDiagrams.length} Mermaid diagrams were converted to images`);
        }
        return warnings;
    }
    /**
     * Cleanup resources
     */
    async cleanup() {
        await this.mermaidProcessor.cleanup();
        this.linkProcessor.cleanup();
        this.processedDiagrams = [];
        this.bookmarkCounter = 0;
    }
}
exports.MarkdownToDocxConverter = MarkdownToDocxConverter;
//# sourceMappingURL=markdown-to-docx-backup.js.map