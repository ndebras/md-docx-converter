"use strict";
/**
 * Main entry point for the Markdown-DOCX converter
 */
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
var __exportStar = (this && this.__exportStar) || function(m, exports) {
    for (var p in m) if (p !== "default" && !Object.prototype.hasOwnProperty.call(exports, p)) __createBinding(exports, m, p);
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MarkdownDocxConverter = void 0;
__exportStar(require("./types"), exports);
__exportStar(require("./converters/markdown-to-docx"), exports);
__exportStar(require("./converters/docx-to-markdown"), exports);
__exportStar(require("./processors/mermaid-processor"), exports);
__exportStar(require("./processors/link-processor"), exports);
__exportStar(require("./styles/document-templates"), exports);
__exportStar(require("./utils"), exports);
const markdown_to_docx_1 = require("./converters/markdown-to-docx");
const docx_to_markdown_1 = require("./converters/docx-to-markdown");
const utils_1 = require("./utils");
const logger_1 = require("./utils/logger");
/**
 * Main converter class providing simplified API
 */
class MarkdownDocxConverter {
    constructor(defaultOptions = {}) {
        this.defaultOptions = defaultOptions;
        this.markdownToDocxConverter = null;
        this.docxToMarkdownConverter = null;
    }
    /**
     * Convert Markdown content to DOCX buffer
     */
    async markdownToDocx(markdownContent, options = {}) {
        const mergedOptions = { ...this.defaultOptions, ...options };
        this.markdownToDocxConverter = new markdown_to_docx_1.MarkdownToDocxConverter();
        try {
            return await this.markdownToDocxConverter.convert(markdownContent, mergedOptions);
        }
        finally {
            this.markdownToDocxConverter = null;
        }
    }
    /**
     * Convert Markdown file to DOCX file
     */
    async markdownFileToDocx(inputFilePath, outputFilePath, options = {}) {
        try {
            logger_1.logger.info('Converting Markdown file to DOCX', { input: inputFilePath, output: outputFilePath });
            // Validate input file
            const validation = utils_1.FileUtils.validateFile(inputFilePath, ['.md', '.markdown']);
            if (!validation.isValid) {
                throw new Error(`Invalid input file: ${validation.errors.map(e => e.message).join(', ')}`);
            }
            // Read markdown content
            const markdownContent = await utils_1.FileUtils.readFile(inputFilePath);
            // Convert to DOCX
            const result = await this.markdownToDocx(markdownContent, options);
            if (result.success && result.output) {
                // Save DOCX file
                await utils_1.FileUtils.writeFile(outputFilePath, result.output);
                logger_1.logger.info('DOCX file saved successfully', { path: outputFilePath });
            }
            return result;
        }
        catch (error) {
            logger_1.logger.error('Failed to convert Markdown file to DOCX', { error, input: inputFilePath });
            return {
                success: false,
                error: {
                    code: 'FILE_CONVERSION_FAILED',
                    message: error instanceof Error ? error.message : String(error),
                    details: { input: inputFilePath, output: outputFilePath },
                },
            };
        }
    }
    /**
     * Convert DOCX file to Markdown content
     */
    async docxToMarkdown(docxFilePath, options = {}) {
        const mergedOptions = { ...this.defaultOptions, ...options };
        this.docxToMarkdownConverter = new docx_to_markdown_1.DocxToMarkdownConverter(mergedOptions);
        try {
            return await this.docxToMarkdownConverter.convert(docxFilePath, mergedOptions);
        }
        finally {
            this.docxToMarkdownConverter = null;
        }
    }
    /**
     * Convert DOCX file to Markdown file
     */
    async docxFileToMarkdown(inputFilePath, outputFilePath, options = {}) {
        try {
            logger_1.logger.info('Converting DOCX file to Markdown', { input: inputFilePath, output: outputFilePath });
            // Convert to Markdown
            const result = await this.docxToMarkdown(inputFilePath, options);
            if (result.success && result.output) {
                // Save Markdown file
                await utils_1.FileUtils.writeFile(outputFilePath, result.output);
                logger_1.logger.info('Markdown file saved successfully', { path: outputFilePath });
            }
            return result;
        }
        catch (error) {
            logger_1.logger.error('Failed to convert DOCX file to Markdown', { error, input: inputFilePath });
            return {
                success: false,
                error: {
                    code: 'FILE_CONVERSION_FAILED',
                    message: error instanceof Error ? error.message : String(error),
                    details: { input: inputFilePath, output: outputFilePath },
                },
            };
        }
    }
    /**
     * Batch convert multiple Markdown files to DOCX
     */
    async batchMarkdownToDocx(inputFiles, outputDir, options = {}) {
        logger_1.logger.info(`Starting batch conversion of ${inputFiles.length} Markdown files`);
        await utils_1.FileUtils.ensureWritableDir(outputDir);
        const results = [];
        for (const inputFile of inputFiles) {
            const fileName = utils_1.FileUtils.getRelativePath(inputFile, inputFile).replace(/\.(md|markdown)$/i, '.docx');
            const outputFile = `${outputDir}/${fileName}`;
            try {
                const result = await this.markdownFileToDocx(inputFile, outputFile, options);
                results.push({ input: inputFile, output: outputFile, result });
            }
            catch (error) {
                logger_1.logger.error(`Failed to convert file: ${inputFile}`, { error });
                results.push({
                    input: inputFile,
                    output: outputFile,
                    result: {
                        success: false,
                        error: {
                            code: 'BATCH_CONVERSION_ERROR',
                            message: error instanceof Error ? error.message : String(error),
                        },
                    },
                });
            }
        }
        const successful = results.filter(r => r.result.success).length;
        logger_1.logger.info(`Batch conversion completed: ${successful}/${inputFiles.length} files converted successfully`);
        return results;
    }
    /**
     * Batch convert multiple DOCX files to Markdown
     */
    async batchDocxToMarkdown(inputFiles, outputDir, options = {}) {
        logger_1.logger.info(`Starting batch conversion of ${inputFiles.length} DOCX files`);
        await utils_1.FileUtils.ensureWritableDir(outputDir);
        const results = [];
        for (const inputFile of inputFiles) {
            const fileName = utils_1.FileUtils.getRelativePath(inputFile, inputFile).replace(/\.docx$/i, '.md');
            const outputFile = `${outputDir}/${fileName}`;
            try {
                const result = await this.docxFileToMarkdown(inputFile, outputFile, options);
                results.push({ input: inputFile, output: outputFile, result });
            }
            catch (error) {
                logger_1.logger.error(`Failed to convert file: ${inputFile}`, { error });
                results.push({
                    input: inputFile,
                    output: outputFile,
                    result: {
                        success: false,
                        error: {
                            code: 'BATCH_CONVERSION_ERROR',
                            message: error instanceof Error ? error.message : String(error),
                        },
                    },
                });
            }
        }
        const successful = results.filter(r => r.result.success).length;
        logger_1.logger.info(`Batch conversion completed: ${successful}/${inputFiles.length} files converted successfully`);
        return results;
    }
    /**
     * Get conversion statistics
     */
    async getConversionStats(filePath) {
        try {
            const content = await utils_1.FileUtils.readFile(filePath);
            const stats = await utils_1.PerformanceUtils.measureAsync(async () => {
                const wordCount = content.split(/\s+/).length;
                const readingTime = Math.ceil(wordCount / 200); // 200 WPM
                const lineCount = content.split('\n').length;
                const imageCount = (content.match(/!\[.*?\]\(.*?\)/g) || []).length;
                const linkCount = (content.match(/\[.*?\]\(.*?\)/g) || []).length;
                return {
                    wordCount,
                    readingTime,
                    lineCount,
                    imageCount,
                    linkCount,
                };
            });
            const fileSize = utils_1.PerformanceUtils.formatBytes(Buffer.from(content, 'utf-8').length);
            return {
                fileSize,
                ...stats.result,
            };
        }
        catch (error) {
            logger_1.logger.error('Failed to get conversion stats', { error, file: filePath });
            throw error;
        }
    }
    /**
     * Validate Markdown content for conversion
     */
    async validateMarkdown(content) {
        const warnings = [];
        const suggestions = [];
        try {
            // Check for common issues
            const lines = content.split('\n');
            // Check for proper heading structure
            const headings = lines.filter(line => line.match(/^#{1,6}\s/));
            if (headings.length === 0) {
                warnings.push('No headings found - consider adding section headers');
            }
            // Check for broken links
            const links = content.match(/\[.*?\]\(.*?\)/g) || [];
            for (const link of links) {
                const urlMatch = link.match(/\((.*?)\)/);
                if (urlMatch && urlMatch[1]) {
                    const url = urlMatch[1];
                    if (url.startsWith('http') && !url.includes('://')) {
                        warnings.push(`Potentially malformed URL: ${url}`);
                    }
                }
            }
            // Check for Mermaid diagrams
            const mermaidBlocks = content.match(/```mermaid[\s\S]*?```/g) || [];
            if (mermaidBlocks.length > 0) {
                suggestions.push(`Found ${mermaidBlocks.length} Mermaid diagram(s) - will be converted to images`);
            }
            // Check for tables
            const tables = content.match(/\|.*\|/g) || [];
            if (tables.length > 0) {
                suggestions.push(`Found ${tables.length} table row(s) - formatting will be preserved`);
            }
            return {
                isValid: warnings.length === 0,
                warnings,
                suggestions,
            };
        }
        catch (error) {
            logger_1.logger.error('Failed to validate Markdown', { error });
            return {
                isValid: false,
                warnings: ['Failed to validate Markdown content'],
                suggestions: [],
            };
        }
    }
    /**
     * Get available templates
     */
    getAvailableTemplates() {
        return [
            'simple',
            'professional-report',
            'technical-documentation',
            'business-proposal',
            'academic-paper',
            'modern',
            'classic',
        ];
    }
    /**
     * Get available Mermaid themes
     */
    getAvailableMermaidThemes() {
        return ['default', 'forest', 'dark', 'neutral', 'base'];
    }
}
exports.MarkdownDocxConverter = MarkdownDocxConverter;
// Default export for convenience
exports.default = MarkdownDocxConverter;
//# sourceMappingURL=index.js.map