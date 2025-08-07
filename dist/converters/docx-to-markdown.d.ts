import { DocxToMarkdownOptions, ConversionResult } from '../types';
/**
 * Advanced DOCX to Markdown converter with formatting preservation
 */
export declare class DocxToMarkdownConverter {
    private turndownService;
    private extractedImages;
    constructor(options?: DocxToMarkdownOptions);
    /**
     * Convert DOCX to Markdown
     */
    convert(docxFilePath: string, options?: DocxToMarkdownOptions): Promise<ConversionResult>;
    /**
     * Configure Turndown service
     */
    private configureTurndown;
    /**
     * Add custom Turndown rules
     */
    private addCustomTurndownRules;
    /**
     * Configure mammoth options
     */
    private configureMammoth;
    /**
     * Extract DOCX content
     */
    private extractDocxContent;
    /**
     * Handle image extraction
     */
    private handleImageExtraction;
    /**
     * Get image file extension from content type
     */
    private getImageExtension;
    /**
     * Convert HTML to Markdown
     */
    private convertHtmlToMarkdown;
    /**
     * Clean HTML for better Markdown conversion
     */
    private cleanHtml;
    /**
     * Fix list formatting in HTML
     */
    private fixListFormatting;
    /**
     * Fix table formatting in HTML
     */
    private fixTableFormatting;
    /**
     * Convert HTML table to Markdown
     */
    private convertTableToMarkdown;
    /**
     * Detect code language from HTML element
     */
    private detectCodeLanguage;
    /**
     * Post-process markdown
     */
    private postProcessMarkdown;
    /**
     * Fix common Markdown formatting issues
     */
    private fixMarkdownFormatting;
    /**
     * Save extracted images to filesystem
     */
    private saveExtractedImages;
    /**
     * Get conversion warnings
     */
    private getConversionWarnings;
    /**
     * Cleanup resources
     */
    private cleanup;
}
//# sourceMappingURL=docx-to-markdown.d.ts.map