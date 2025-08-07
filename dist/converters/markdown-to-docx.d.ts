import type { MarkdownToDocxOptions, ConversionResult } from '../types';
export declare class MarkdownToDocxConverter {
    private mermaidProcessor;
    private linkProcessor;
    private logger;
    private bookmarkCounter;
    private headings;
    private mermaidImages;
    constructor();
    convert(markdownContent: string, options?: MarkdownToDocxOptions): Promise<ConversionResult>;
    private processTokens;
    private processToken;
    private createHeading;
    private createParagraph;
    private createImageParagraph;
    private createList;
    private createBlockquote;
    private createCodeBlock;
    private createTable;
    private createHtml;
    private createHorizontalRule;
    private processInlineTokens;
    private processInlineTokensWithStyle;
    private parseBasicInlineFormatting;
    private createTableOfContents;
    private extractHeadings;
    private createTableOfContentsWithLinks;
    private createDocument;
    private generateAnchor;
    private extractTextFromTokens;
    private getHeadingLevel;
    private countWords;
    private decodeHtmlEntities;
    private parseSimpleMarkdown;
}
//# sourceMappingURL=markdown-to-docx.d.ts.map