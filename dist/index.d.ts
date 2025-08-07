/**
 * Main entry point for the Markdown-DOCX converter
 */
export * from './types';
export * from './converters/markdown-to-docx';
export * from './converters/docx-to-markdown';
export * from './processors/mermaid-processor';
export * from './processors/link-processor';
export * from './styles/document-templates';
export * from './utils';
import { MarkdownToDocxOptions, DocxToMarkdownOptions, ConversionResult, DocumentTemplate, MermaidTheme } from './types';
/**
 * Main converter class providing simplified API
 */
export declare class MarkdownDocxConverter {
    private defaultOptions;
    private markdownToDocxConverter;
    private docxToMarkdownConverter;
    constructor(defaultOptions?: {
        template?: DocumentTemplate;
        mermaidTheme?: MermaidTheme;
        preserveLinks?: boolean;
        tocGeneration?: boolean;
    });
    /**
     * Convert Markdown content to DOCX buffer
     */
    markdownToDocx(markdownContent: string, options?: MarkdownToDocxOptions): Promise<ConversionResult>;
    /**
     * Convert Markdown file to DOCX file
     */
    markdownFileToDocx(inputFilePath: string, outputFilePath: string, options?: MarkdownToDocxOptions): Promise<ConversionResult>;
    /**
     * Convert DOCX file to Markdown content
     */
    docxToMarkdown(docxFilePath: string, options?: DocxToMarkdownOptions): Promise<ConversionResult>;
    /**
     * Convert DOCX file to Markdown file
     */
    docxFileToMarkdown(inputFilePath: string, outputFilePath: string, options?: DocxToMarkdownOptions): Promise<ConversionResult>;
    /**
     * Batch convert multiple Markdown files to DOCX
     */
    batchMarkdownToDocx(inputFiles: string[], outputDir: string, options?: MarkdownToDocxOptions): Promise<Array<{
        input: string;
        output: string;
        result: ConversionResult;
    }>>;
    /**
     * Batch convert multiple DOCX files to Markdown
     */
    batchDocxToMarkdown(inputFiles: string[], outputDir: string, options?: DocxToMarkdownOptions): Promise<Array<{
        input: string;
        output: string;
        result: ConversionResult;
    }>>;
    /**
     * Get conversion statistics
     */
    getConversionStats(filePath: string): Promise<{
        fileSize: string;
        wordCount: number;
        readingTime: number;
        lineCount: number;
        imageCount: number;
        linkCount: number;
    }>;
    /**
     * Validate Markdown content for conversion
     */
    validateMarkdown(content: string): Promise<{
        isValid: boolean;
        warnings: string[];
        suggestions: string[];
    }>;
    /**
     * Get available templates
     */
    getAvailableTemplates(): DocumentTemplate[];
    /**
     * Get available Mermaid themes
     */
    getAvailableMermaidThemes(): MermaidTheme[];
}
export default MarkdownDocxConverter;
//# sourceMappingURL=index.d.ts.map