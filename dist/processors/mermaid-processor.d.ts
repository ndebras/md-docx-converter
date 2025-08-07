import { ProcessedMermaidDiagram, MermaidTheme } from '../types';
/**
 * Advanced Mermaid diagram processor with multiple output formats
 */
export declare class MermaidProcessor {
    private theme;
    private browser;
    private page;
    private isInitialized;
    constructor(theme?: MermaidTheme);
    /**
     * Process content containing Mermaid diagrams
     */
    processContent(content: string, options?: {
        theme?: MermaidTheme;
    }): Promise<{
        content: string;
        diagramCount: number;
    }>;
    /**
     * Create an enhanced textual representation of Mermaid diagrams for DOCX
     */
    private createEnhancedTextualDiagramRepresentation;
    /**
     * Create enhanced flowchart text representation with better parsing
     */
    private createEnhancedFlowchartText;
    /**
     * Create enhanced sequence diagram text representation
     */
    private createEnhancedSequenceText;
    /**
     * Create enhanced class diagram text representation
     */
    private createEnhancedClassDiagramText;
    /**
     * Create enhanced generic diagram text representation
     */
    private createEnhancedGenericDiagramText;
    /**
     * Create a textual representation of Mermaid diagrams for DOCX
     */
    private createTextualDiagramRepresentation;
    /**
     * Create flowchart text representation
     */
    private createFlowchartText;
    /**
     * Extract node information (ID and label)
     */
    private extractNodeInfo;
    /**
     * Create sequence diagram text representation
     */
    private createSequenceText;
    /**
     * Create class diagram text representation
     */
    private createClassDiagramText;
    /**
     * Create generic diagram text representation
     */
    private createGenericDiagramText;
    /**
     * Generate a simple SVG placeholder for Mermaid diagrams
     */
    private generateMermaidSVGPlaceholder;
    /**
     * Generate a flowchart-style SVG
     */
    private generateFlowchartSVG;
    /**
     * Generate a sequence diagram SVG
     */
    private generateSequenceSVG;
    /**
     * Generate a default diagram SVG
     */
    private generateDefaultDiagramSVG;
    /**
     * Detect the type of Mermaid diagram
     */
    private detectDiagramType;
    /**
     * Initialize Puppeteer browser
     */
    initialize(): Promise<void>;
    /**
     * Process Mermaid diagram and return SVG/PNG
     */
    processDiagram(mermaidCode: string, format?: 'svg' | 'png', options?: {
        width?: number;
        height?: number;
        backgroundColor?: string;
        scale?: number;
    }): Promise<ProcessedMermaidDiagram>;
    /**
     * Extract all Mermaid diagrams from Markdown content
     */
    extractMermaidDiagrams(markdown: string): Array<{
        code: string;
        startLine: number;
        endLine: number;
        id: string;
    }>;
    /**
     * Replace Mermaid code blocks with image references
     */
    replaceMermaidBlocks(markdown: string, processedDiagrams: ProcessedMermaidDiagram[], imageBasePath?: string): string;
    /**
     * Validate Mermaid diagram syntax
     */
    validateDiagram(mermaidCode: string): Promise<{
        isValid: boolean;
        error?: string;
    }>;
    /**
     * Get supported diagram types
     */
    getSupportedTypes(): string[];
    /**
     * Cleanup resources
     */
    cleanup(): Promise<void>;
    /**
     * Private method to render diagram
     */
    private renderDiagram;
}
/**
 * Mermaid theme configurations
 */
export declare const MermaidThemes: Record<MermaidTheme, Record<string, any>>;
//# sourceMappingURL=mermaid-processor.d.ts.map