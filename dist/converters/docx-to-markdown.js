"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.DocxToMarkdownConverter = void 0;
const mammoth = __importStar(require("mammoth"));
const turndown_1 = __importDefault(require("turndown"));
const fs = __importStar(require("fs-extra"));
const path = __importStar(require("path"));
const utils_1 = require("../utils");
const logger_1 = require("../utils/logger");
/**
 * Advanced DOCX to Markdown converter with formatting preservation
 */
class DocxToMarkdownConverter {
    constructor(options = {}) {
        this.extractedImages = [];
        this.turndownService = this.configureTurndown(options);
    }
    /**
     * Convert DOCX to Markdown
     */
    async convert(docxFilePath, options = {}) {
        const startTime = Date.now();
        try {
            logger_1.logger.info('Starting DOCX to Markdown conversion', { file: docxFilePath });
            // Validate input file
            const validation = utils_1.FileUtils.validateFile(docxFilePath, ['.docx']);
            if (!validation.isValid) {
                throw new Error(`Invalid input file: ${validation.errors.map(e => e.message).join(', ')}`);
            }
            // Read DOCX file
            const { result: docxBuffer, elapsed: readTime } = await utils_1.PerformanceUtils.measureAsync(async () => {
                return await fs.readFile(docxFilePath);
            }, 'file-read');
            logger_1.logger.debug(`File read completed in ${readTime}ms`);
            // Configure mammoth options
            const mammothOptions = this.configureMammoth(options);
            // Extract content and images
            const { result: extraction, elapsed: extractTime } = await utils_1.PerformanceUtils.measureAsync(async () => {
                return await this.extractDocxContent(docxBuffer, mammothOptions, options);
            }, 'content-extraction');
            logger_1.logger.debug(`Content extraction completed in ${extractTime}ms`);
            // Convert HTML to Markdown
            const { result: markdown, elapsed: conversionTime } = await utils_1.PerformanceUtils.measureAsync(async () => {
                return this.convertHtmlToMarkdown(extraction.html, options);
            }, 'html-to-markdown');
            logger_1.logger.debug(`HTML to Markdown conversion completed in ${conversionTime}ms`);
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
            logger_1.logger.info('DOCX to Markdown conversion completed successfully', metadata);
            return {
                success: true,
                output: processedMarkdown,
                metadata,
                warnings: this.getConversionWarnings(extraction.messages),
            };
        }
        catch (error) {
            logger_1.logger.error('DOCX to Markdown conversion failed', { error, file: docxFilePath });
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
        }
        finally {
            this.cleanup();
        }
    }
    /**
     * Configure Turndown service
     */
    configureTurndown(options) {
        const turndown = new turndown_1.default({
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
    addCustomTurndownRules(turndown, options) {
        // Preserve code blocks
        turndown.addRule('codeBlock', {
            filter: ['pre'],
            replacement: (content, node) => {
                const language = this.detectCodeLanguage(node);
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
                return this.convertTableToMarkdown(node);
            },
        });
        // Handle images
        turndown.addRule('image', {
            filter: 'img',
            replacement: (content, node) => {
                const img = node;
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
                return node.nodeName === 'A' && node.href ? true : false;
            },
            replacement: (content, node) => {
                const link = node;
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
    configureMammoth(options) {
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
            convertImage: (image) => this.handleImageExtraction(image, options),
            ignoreEmptyParagraphs: false,
            preserveLineBreaks: options.preserveFormatting || false,
        };
    }
    /**
     * Extract DOCX content
     */
    async extractDocxContent(docxBuffer, mammothOptions, options) {
        try {
            const result = await mammoth.convertToHtml(docxBuffer, mammothOptions);
            logger_1.logger.debug('DOCX extraction completed', {
                messageCount: result.messages.length,
                warnings: result.messages.filter((m) => m.type === 'warning').length,
                errors: result.messages.filter((m) => m.type === 'error').length,
            });
            return {
                html: result.value,
                messages: result.messages,
            };
        }
        catch (error) {
            logger_1.logger.error('Failed to extract DOCX content', { error });
            throw error;
        }
    }
    /**
     * Handle image extraction
     */
    handleImageExtraction(image, options) {
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
            logger_1.logger.debug(`Extracted image: ${filename}`);
            return {
                src: filename,
                altText: image.altText || 'Extracted image',
            };
        }
        catch (error) {
            logger_1.logger.warn('Failed to extract image', { error });
            return image;
        }
    }
    /**
     * Get image file extension from content type
     */
    getImageExtension(contentType) {
        const mimeToExt = {
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
    convertHtmlToMarkdown(html, options) {
        try {
            // Clean up HTML before conversion
            const cleanedHtml = this.cleanHtml(html, options);
            // Convert to Markdown
            const markdown = this.turndownService.turndown(cleanedHtml);
            return markdown;
        }
        catch (error) {
            logger_1.logger.error('Failed to convert HTML to Markdown', { error });
            throw error;
        }
    }
    /**
     * Clean HTML for better Markdown conversion
     */
    cleanHtml(html, options) {
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
    fixListFormatting(html) {
        // Convert Word list paragraphs to proper HTML lists
        return html.replace(/<p[^>]*>\s*([?????\-\*]|\d+\.)\s*([^<]+)<\/p>/g, '<li>$2</li>');
    }
    /**
     * Fix table formatting in HTML
     */
    fixTableFormatting(html) {
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
    convertTableToMarkdown(table) {
        const rows = [];
        // Extract table data
        const tableRows = table.querySelectorAll('tr');
        for (let i = 0; i < tableRows.length; i++) {
            const row = tableRows[i];
            const cells = [];
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
        const markdownRows = [];
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
    detectCodeLanguage(element) {
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
    postProcessMarkdown(markdown, options) {
        let processed = markdown;
        // Normalize whitespace
        processed = utils_1.StringUtils.normalizeWhitespace(processed);
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
    fixMarkdownFormatting(markdown) {
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
    async saveExtractedImages(outputDir) {
        if (this.extractedImages.length === 0) {
            return;
        }
        await utils_1.FileUtils.ensureWritableDir(outputDir);
        for (const image of this.extractedImages) {
            const imagePath = path.join(outputDir, image.filename);
            await fs.writeFile(imagePath, image.buffer);
            logger_1.logger.debug(`Saved extracted image: ${imagePath}`);
        }
        logger_1.logger.info(`Saved ${this.extractedImages.length} extracted images to ${outputDir}`);
    }
    /**
     * Get conversion warnings
     */
    getConversionWarnings(messages) {
        const warnings = [];
        const messagesByType = messages.reduce((acc, msg) => {
            acc[msg.type] = (acc[msg.type] || 0) + 1;
            return acc;
        }, {});
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
    cleanup() {
        this.extractedImages = [];
    }
}
exports.DocxToMarkdownConverter = DocxToMarkdownConverter;
//# sourceMappingURL=docx-to-markdown.js.map