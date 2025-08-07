/**
 * Core types and interfaces for the Markdown-DOCX converter
 */
export interface ConversionOptions {
    /** Template to use for document styling */
    template?: DocumentTemplate;
    /** Theme for Mermaid diagrams */
    mermaidTheme?: MermaidTheme;
    /** Whether to preserve internal links */
    preserveLinks?: boolean;
    /** Whether to generate table of contents */
    tocGeneration?: boolean;
    /** Whether to include metadata in document */
    includeMetadata?: boolean;
    /** Custom styles to apply */
    customStyles?: CustomStyles;
    /** Output format options */
    outputOptions?: OutputOptions;
}
export interface MarkdownToDocxOptions extends ConversionOptions {
    /** Document title */
    title?: string;
    /** Document author */
    author?: string;
    /** Document subject */
    subject?: string;
    /** Document description */
    description?: string;
    /** Document keywords */
    keywords?: string[];
    /** Document creation date */
    createdAt?: Date;
    /** Page orientation */
    orientation?: 'portrait' | 'landscape';
    /** Page margins */
    margins?: PageMargins;
}
export interface DocxToMarkdownOptions extends ConversionOptions {
    /** Whether to preserve original formatting */
    preserveFormatting?: boolean;
    /** Whether to extract images to separate files */
    extractImages?: boolean;
    /** Output directory for extracted images */
    imageOutputDir?: string;
    /** Image format for extraction */
    imageFormat?: 'png' | 'jpg' | 'svg';
}
export interface PageMargins {
    top: number;
    right: number;
    bottom: number;
    left: number;
}
export interface CustomStyles {
    /** Heading styles */
    headings?: HeadingStyles;
    /** Paragraph styles */
    paragraph?: ParagraphStyle;
    /** Code block styles */
    codeBlock?: CodeBlockStyle;
    /** Table styles */
    table?: TableStyle;
    /** Link styles */
    link?: LinkStyle;
}
export interface HeadingStyles {
    h1?: HeadingStyle;
    h2?: HeadingStyle;
    h3?: HeadingStyle;
    h4?: HeadingStyle;
    h5?: HeadingStyle;
    h6?: HeadingStyle;
}
export interface HeadingStyle {
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    alignment?: 'left' | 'center' | 'right' | 'justify';
    spacing?: {
        before?: number;
        after?: number;
    };
}
export interface ParagraphStyle {
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    alignment?: 'left' | 'center' | 'right' | 'justify';
    spacing?: {
        before?: number;
        after?: number;
        lineSpacing?: number;
    };
    indent?: {
        left?: number;
        right?: number;
        firstLine?: number;
    };
}
export interface CodeBlockStyle {
    fontSize?: number;
    fontFamily?: string;
    backgroundColor?: string;
    textColor?: string;
    border?: {
        color?: string;
        width?: number;
        style?: 'solid' | 'dashed' | 'dotted';
    };
    padding?: number;
}
export interface TableStyle {
    headerStyle?: {
        backgroundColor?: string;
        textColor?: string;
        bold?: boolean;
    };
    borderStyle?: {
        color?: string;
        width?: number;
        style?: 'solid' | 'dashed' | 'dotted';
    };
    cellPadding?: number;
    alternateRowShading?: boolean;
}
export interface LinkStyle {
    color?: string;
    underline?: boolean;
    bold?: boolean;
    italic?: boolean;
}
export interface OutputOptions {
    /** Whether to compress output */
    compress?: boolean;
    /** Quality settings for images */
    imageQuality?: number;
    /** Whether to optimize for file size */
    optimizeSize?: boolean;
}
export type DocumentTemplate = 'professional-report' | 'technical-documentation' | 'business-proposal' | 'academic-paper' | 'simple' | 'modern' | 'classic';
export type MermaidTheme = 'default' | 'forest' | 'dark' | 'neutral' | 'base';
export interface ConversionResult {
    /** Success status */
    success: boolean;
    /** Output buffer or file path */
    output?: Buffer | string;
    /** Conversion metadata */
    metadata?: ConversionMetadata;
    /** Any warnings during conversion */
    warnings?: string[];
    /** Error information if conversion failed */
    error?: ConversionError;
}
export interface ConversionMetadata {
    /** Input file size */
    inputSize: number;
    /** Output file size */
    outputSize: number;
    /** Processing time in milliseconds */
    processingTime: number;
    /** Number of pages generated */
    pageCount?: number;
    /** Number of images processed */
    imageCount?: number;
    /** Number of Mermaid diagrams converted */
    mermaidDiagramCount?: number;
    /** Number of internal links processed */
    internalLinkCount?: number;
    /** Number of external links processed */
    externalLinkCount?: number;
}
export interface ConversionError {
    /** Error code */
    code: string;
    /** Error message */
    message: string;
    /** Detailed error information */
    details?: Record<string, unknown>;
    /** Stack trace */
    stack?: string;
}
export interface ProcessedMermaidDiagram {
    /** Original Mermaid code */
    code: string;
    /** Generated image buffer */
    imageBuffer: Buffer;
    /** Image format */
    format: 'svg' | 'png';
    /** Image dimensions */
    dimensions: {
        width: number;
        height: number;
    };
    /** Unique identifier */
    id: string;
}
export interface ProcessedLink {
    /** Original link text */
    text: string;
    /** Link URL or anchor */
    url: string;
    /** Link type */
    type: 'internal' | 'external' | 'anchor';
    /** Whether link is valid */
    isValid: boolean;
    /** Processed URL for document */
    processedUrl?: string;
}
export interface DocumentSection {
    /** Section title */
    title: string;
    /** Section level (1-6) */
    level: number;
    /** Section content */
    content: string;
    /** Unique identifier */
    id: string;
    /** Child sections */
    children: DocumentSection[];
}
export interface TableOfContents {
    /** TOC entries */
    entries: TOCEntry[];
    /** TOC title */
    title: string;
    /** Whether to include page numbers */
    includePageNumbers: boolean;
}
export interface TOCEntry {
    /** Entry title */
    title: string;
    /** Entry level */
    level: number;
    /** Page number */
    pageNumber?: number;
    /** Anchor link */
    anchor: string;
    /** Child entries */
    children: TOCEntry[];
}
export interface ValidationResult {
    /** Whether validation passed */
    isValid: boolean;
    /** Validation errors */
    errors: ValidationError[];
    /** Validation warnings */
    warnings: ValidationWarning[];
}
export interface ValidationError {
    /** Error code */
    code: string;
    /** Error message */
    message: string;
    /** Line number where error occurred */
    line?: number;
    /** Column number where error occurred */
    column?: number;
}
export interface ValidationWarning {
    /** Warning code */
    code: string;
    /** Warning message */
    message: string;
    /** Line number where warning occurred */
    line?: number;
    /** Column number where warning occurred */
    column?: number;
}
export interface LoggerOptions {
    /** Log level */
    level: 'debug' | 'info' | 'warn' | 'error';
    /** Whether to include timestamps */
    timestamp: boolean;
    /** Whether to use colors */
    colors: boolean;
    /** Output stream */
    stream?: NodeJS.WriteStream;
}
export interface ProgressCallback {
    (progress: ProgressInfo): void;
}
export interface ProgressInfo {
    /** Current step */
    step: string;
    /** Progress percentage (0-100) */
    percentage: number;
    /** Current item being processed */
    currentItem?: string;
    /** Total items to process */
    totalItems?: number;
    /** Completed items */
    completedItems?: number;
}
//# sourceMappingURL=index.d.ts.map