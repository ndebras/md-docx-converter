import { DocumentTemplate } from '../types';
/**
 * Professional document templates for DOCX generation
 */
export declare class DocumentTemplates {
    /**
     * Get template configuration
     */
    static getTemplate(templateName: DocumentTemplate): DocumentTemplateConfig;
    /**
     * Professional report template
     */
    private static createProfessionalReportTemplate;
    /**
     * Technical documentation template
     */
    private static createTechnicalDocumentationTemplate;
    /**
     * Business proposal template
     */
    private static createBusinessProposalTemplate;
    /**
     * Academic paper template
     */
    private static createAcademicPaperTemplate;
    /**
     * Modern template
     */
    private static createModernTemplate;
    /**
     * Classic template
     */
    private static createClassicTemplate;
    /**
     * Simple template
     */
    private static createSimpleTemplate;
}
/**
 * Document template configuration interface
 */
export interface DocumentTemplateConfig {
    styles: {
        default: any;
        headings: Record<string, any>;
        paragraph?: any;
        codeBlock?: any;
        table?: any;
        hyperlink?: any;
    };
    pageSettings: {
        margins: {
            top: number;
            right: number;
            bottom: number;
            left: number;
        };
        orientation: 'portrait' | 'landscape';
    };
    sections: {
        properties: any;
        headers?: any;
        footers?: any;
    };
}
//# sourceMappingURL=document-templates.d.ts.map