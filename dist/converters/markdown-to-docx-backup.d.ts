import { MarkdownToDocxOptions, ConversionResult } from '../types';
/**
 * Advanced Markdown to DOCX converter with professional styling
 */
export declare class MarkdownToDocxConverter {
    private mermaidProcessor;
    private linkProcessor;
    private templateConfig;
    private processedDiagrams;
    private bookmarkCounter;
    constructor(options?: MarkdownToDocxOptions);
    /**
     * Convert Markdown to DOCX
     */
    convert(markdownContent: string, options?: MarkdownToDocxOptions): Promise<ConversionResult>;
    /**
     * Process Mermaid diagrams
     */
    private processMermaidDiagrams;
    /**
     * Process markdown tokens to DOCX elements
     */
    private processTokens;
    /**
     * Process individual markdown token
     */
    private processToken;
    /**
     * Create heading element
     */
    private createHeading;
    /**
     * Create paragraph element
     */
    private createParagraph;
    /**
     * Create list elements
     */
    private createList;
    /**
     * Create blockquote element
     */
    private createBlockquote;
    /**
     * Create code block element
     */
    private createCodeBlock;
    /**
     * Create table element
     */
    private createTable;
    /**
     * Create horizontal rule
     */
    private createHorizontalRule;
    /**
     * Create HTML element (limited support)
     */
    private createHtml;
    /**
     * Process inline tokens (bold, italic, links, etc.)
     */
    private processInlineTokens;
    /**
     * Extract plain text from tokens
     */
    private extractTextFromTokens;
    /**
     * Get heading style based on depth
     */
    private getHeadingStyle;
    /**
     * Get DOCX heading level
     */
    private getHeadingLevel;
    /**
     * Create final document
     */
    private createDocument;
    /**
     * Get conversion warnings
     */
    private getConversionWarnings;
    /**
     * Cleanup resources
     */
    private cleanup;
}
//# sourceMappingURL=markdown-to-docx-backup.d.ts.map