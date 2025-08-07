import { ProcessedLink } from '../types';
/**
 * Advanced link processor for Markdown and DOCX conversion
 */
export declare class LinkProcessor {
    private anchors;
    private externalLinks;
    private internalLinks;
    /**
     * Process content containing links
     */
    processContent(content: string): {
        content: string;
        links: any[];
        headings: any[];
    };
    /**
     * Process all links in Markdown content
     */
    processMarkdownLinks(markdown: string): {
        processedContent: string;
        links: ProcessedLink[];
    };
    /**
     * Convert Markdown links to Word hyperlinks
     */
    convertToWordLinks(links: ProcessedLink[]): Array<{
        text: string;
        url: string;
        type: 'hyperlink' | 'bookmark';
        style?: string;
    }>;
    /**
     * Extract links from DOCX content
     */
    extractDocxLinks(content: string): ProcessedLink[];
    /**
     * Validate link URLs
     */
    validateLinks(links: ProcessedLink[]): ProcessedLink[];
    /**
     * Generate anchor map from headings
     */
    private buildAnchorMap;
    /**
     * Process internal links [text](#anchor)
     */
    private processInternalLinks;
    /**
     * Process external links [text](http://example.com)
     */
    private processExternalLinks;
    /**
     * Process reference-style links [text][ref]
     */
    private processReferenceLinks;
    /**
     * Process image links ![alt](src)
     */
    private processImageLinks;
    /**
     * Extract reference definitions [ref]: url "title"
     */
    private extractReferenceDefinitions;
    /**
     * Validate HTTP URL
     */
    private isValidHttpUrl;
    /**
     * Create bookmark name for Word document
     */
    private createBookmarkName;
    /**
     * Generate table of contents from headings
     */
    generateTableOfContents(markdown: string): {
        tocMarkdown: string;
        entries: Array<{
            title: string;
            level: number;
            anchor: string;
        }>;
    };
    /**
     * Update internal links after heading changes
     */
    updateInternalLinks(content: string, headingMap: Map<string, string>): string;
    /**
     * Get link statistics
     */
    getLinkStatistics(links: ProcessedLink[]): {
        total: number;
        byType: Record<string, number>;
        valid: number;
        invalid: number;
    };
    /**
     * Clean up resources
     */
    cleanup(): void;
}
//# sourceMappingURL=link-processor.d.ts.map