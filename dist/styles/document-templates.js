"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DocumentTemplates = void 0;
const docx_1 = require("docx");
/**
 * Professional document templates for DOCX generation
 */
class DocumentTemplates {
    /**
     * Get template configuration
     */
    static getTemplate(templateName) {
        switch (templateName) {
            case 'professional-report':
                return this.createProfessionalReportTemplate();
            case 'technical-documentation':
                return this.createTechnicalDocumentationTemplate();
            case 'business-proposal':
                return this.createBusinessProposalTemplate();
            case 'academic-paper':
                return this.createAcademicPaperTemplate();
            case 'modern':
                return this.createModernTemplate();
            case 'classic':
                return this.createClassicTemplate();
            case 'simple':
            default:
                return this.createSimpleTemplate();
        }
    }
    /**
     * Professional report template
     */
    static createProfessionalReportTemplate() {
        return {
            styles: {
                default: {
                    document: {
                        run: {
                            size: 24, // 12pt
                            font: 'Calibri',
                            color: '333333',
                        },
                        paragraph: {
                            spacing: {
                                line: 276, // 1.15 line spacing
                                after: 120,
                            },
                        },
                    },
                },
                headings: {
                    heading1: {
                        run: {
                            size: 32, // 16pt
                            font: 'Calibri Light',
                            color: '2F5597',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 240,
                                after: 120,
                            },
                            border: {
                                bottom: {
                                    color: '2F5597',
                                    space: 1,
                                    style: docx_1.BorderStyle.SINGLE,
                                    size: 6,
                                },
                            },
                        },
                    },
                    heading2: {
                        run: {
                            size: 28, // 14pt
                            font: 'Calibri',
                            color: '2F5597',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 180,
                                after: 60,
                            },
                        },
                    },
                    heading3: {
                        run: {
                            size: 26, // 13pt
                            font: 'Calibri',
                            color: '333333',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 120,
                                after: 60,
                            },
                        },
                    },
                },
                paragraph: {
                    normal: {
                        run: {
                            size: 24,
                            font: 'Calibri',
                        },
                        paragraph: {
                            spacing: {
                                after: 120,
                            },
                            alignment: docx_1.AlignmentType.JUSTIFIED,
                        },
                    },
                },
                codeBlock: {
                    run: {
                        size: 20, // 10pt
                        font: 'Consolas',
                        color: '333333',
                    },
                    paragraph: {
                        shading: {
                            fill: 'F8F8F8',
                        },
                        border: {
                            left: {
                                color: 'CCCCCC',
                                space: 1,
                                style: docx_1.BorderStyle.SINGLE,
                                size: 6,
                            },
                        },
                        indent: {
                            left: 720, // 0.5 inch
                        },
                        spacing: {
                            before: 120,
                            after: 120,
                        },
                    },
                },
                table: {
                    borders: {
                        top: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '2F5597' },
                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '2F5597' },
                        left: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '2F5597' },
                        right: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '2F5597' },
                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
                    },
                    width: {
                        size: 100,
                        type: docx_1.WidthType.PERCENTAGE,
                    },
                },
                hyperlink: {
                    run: {
                        color: '0563C1',
                        underline: {
                            type: docx_1.UnderlineType.SINGLE,
                            color: '0563C1',
                        },
                    },
                },
            },
            pageSettings: {
                margins: {
                    top: 1440, // 1 inch
                    right: 1440,
                    bottom: 1440,
                    left: 1440,
                },
                orientation: 'portrait',
            },
            sections: {
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            right: 1440,
                            bottom: 1440,
                            left: 1440,
                        },
                    },
                },
                headers: {
                    default: new docx_1.Header({
                        children: [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Professional Report",
                                        size: 20,
                                        color: '666666',
                                    }),
                                ],
                                alignment: docx_1.AlignmentType.RIGHT,
                            }),
                        ],
                    }),
                },
                footers: {
                    default: new docx_1.Footer({
                        children: [
                            new docx_1.Paragraph({
                                children: [
                                    new docx_1.TextRun({
                                        text: "Page ",
                                        size: 20,
                                        color: '666666',
                                    }),
                                    new docx_1.TextRun({
                                        children: [docx_1.PageNumber.CURRENT],
                                        size: 20,
                                        color: '666666',
                                    }),
                                    new docx_1.TextRun({
                                        text: " of ",
                                        size: 20,
                                        color: '666666',
                                    }),
                                    new docx_1.TextRun({
                                        children: [docx_1.PageNumber.TOTAL_PAGES],
                                        size: 20,
                                        color: '666666',
                                    }),
                                ],
                                alignment: docx_1.AlignmentType.CENTER,
                            }),
                        ],
                    }),
                },
            },
        };
    }
    /**
     * Technical documentation template
     */
    static createTechnicalDocumentationTemplate() {
        return {
            styles: {
                default: {
                    document: {
                        run: {
                            size: 22, // 11pt
                            font: 'Source Sans Pro',
                            color: '333333',
                        },
                        paragraph: {
                            spacing: {
                                line: 240, // Single line spacing
                                after: 120,
                            },
                        },
                    },
                },
                headings: {
                    heading1: {
                        run: {
                            size: 36, // 18pt
                            font: 'Source Sans Pro',
                            color: '1E88E5',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 360,
                                after: 180,
                            },
                            numbering: {
                                reference: 'heading-numbering',
                                level: 0,
                            },
                        },
                    },
                    heading2: {
                        run: {
                            size: 30, // 15pt
                            font: 'Source Sans Pro',
                            color: '1976D2',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 240,
                                after: 120,
                            },
                            numbering: {
                                reference: 'heading-numbering',
                                level: 1,
                            },
                        },
                    },
                    heading3: {
                        run: {
                            size: 26, // 13pt
                            font: 'Source Sans Pro',
                            color: '424242',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 180,
                                after: 60,
                            },
                            numbering: {
                                reference: 'heading-numbering',
                                level: 2,
                            },
                        },
                    },
                },
                codeBlock: {
                    run: {
                        size: 18, // 9pt
                        font: 'JetBrains Mono',
                        color: '263238',
                    },
                    paragraph: {
                        shading: {
                            fill: 'F5F5F5',
                        },
                        border: {
                            top: { color: 'E0E0E0', space: 1, style: docx_1.BorderStyle.SINGLE, size: 4 },
                            bottom: { color: 'E0E0E0', space: 1, style: docx_1.BorderStyle.SINGLE, size: 4 },
                            left: { color: '1E88E5', space: 1, style: docx_1.BorderStyle.SINGLE, size: 12 },
                            right: { color: 'E0E0E0', space: 1, style: docx_1.BorderStyle.SINGLE, size: 4 },
                        },
                        indent: {
                            left: 360,
                        },
                        spacing: {
                            before: 120,
                            after: 120,
                        },
                    },
                },
                table: {
                    borders: {
                        top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: '1E88E5' },
                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'BDBDBD' },
                        left: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'BDBDBD' },
                        right: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'BDBDBD' },
                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
                    },
                    width: {
                        size: 100,
                        type: docx_1.WidthType.PERCENTAGE,
                    },
                },
            },
            pageSettings: {
                margins: {
                    top: 1440,
                    right: 1080,
                    bottom: 1440,
                    left: 1080,
                },
                orientation: 'portrait',
            },
            sections: {
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            right: 1080,
                            bottom: 1440,
                            left: 1080,
                        },
                    },
                },
            },
        };
    }
    /**
     * Business proposal template
     */
    static createBusinessProposalTemplate() {
        return {
            styles: {
                default: {
                    document: {
                        run: {
                            size: 24, // 12pt
                            font: 'Georgia',
                            color: '333333',
                        },
                        paragraph: {
                            spacing: {
                                line: 360, // 1.5 line spacing
                                after: 240,
                            },
                        },
                    },
                },
                headings: {
                    heading1: {
                        run: {
                            size: 40, // 20pt
                            font: 'Georgia',
                            color: '8B4513',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 480,
                                after: 240,
                            },
                            alignment: docx_1.AlignmentType.CENTER,
                        },
                    },
                    heading2: {
                        run: {
                            size: 32, // 16pt
                            font: 'Georgia',
                            color: '8B4513',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 360,
                                after: 180,
                            },
                        },
                    },
                },
                table: {
                    borders: {
                        top: { style: docx_1.BorderStyle.DOUBLE, size: 4, color: '8B4513' },
                        bottom: { style: docx_1.BorderStyle.DOUBLE, size: 4, color: '8B4513' },
                        left: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '8B4513' },
                        right: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '8B4513' },
                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'DDD' },
                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'DDD' },
                    },
                },
            },
            pageSettings: {
                margins: {
                    top: 1800, // 1.25 inches
                    right: 1440,
                    bottom: 1800,
                    left: 1440,
                },
                orientation: 'portrait',
            },
            sections: {
                properties: {
                    page: {
                        margin: {
                            top: 1800,
                            right: 1440,
                            bottom: 1800,
                            left: 1440,
                        },
                    },
                },
            },
        };
    }
    /**
     * Academic paper template
     */
    static createAcademicPaperTemplate() {
        return {
            styles: {
                default: {
                    document: {
                        run: {
                            size: 24, // 12pt
                            font: 'Times New Roman',
                            color: '000000',
                        },
                        paragraph: {
                            spacing: {
                                line: 480, // Double line spacing
                                after: 0,
                            },
                        },
                    },
                },
                headings: {
                    heading1: {
                        run: {
                            size: 24, // 12pt
                            font: 'Times New Roman',
                            color: '000000',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 240,
                                after: 120,
                            },
                            alignment: docx_1.AlignmentType.CENTER,
                        },
                    },
                    heading2: {
                        run: {
                            size: 24, // 12pt
                            font: 'Times New Roman',
                            color: '000000',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 240,
                                after: 120,
                            },
                        },
                    },
                },
                table: {
                    borders: {
                        top: { style: docx_1.BorderStyle.SINGLE, size: 2, color: '000000' },
                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 2, color: '000000' },
                        left: { style: docx_1.BorderStyle.NONE },
                        right: { style: docx_1.BorderStyle.NONE },
                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '000000' },
                        insideVertical: { style: docx_1.BorderStyle.NONE },
                    },
                },
            },
            pageSettings: {
                margins: {
                    top: 1440, // 1 inch
                    right: 1440,
                    bottom: 1440,
                    left: 1440,
                },
                orientation: 'portrait',
            },
            sections: {
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            right: 1440,
                            bottom: 1440,
                            left: 1440,
                        },
                    },
                },
            },
        };
    }
    /**
     * Modern template
     */
    static createModernTemplate() {
        return {
            styles: {
                default: {
                    document: {
                        run: {
                            size: 22, // 11pt
                            font: 'Segoe UI',
                            color: '333333',
                        },
                        paragraph: {
                            spacing: {
                                line: 276,
                                after: 120,
                            },
                        },
                    },
                },
                headings: {
                    heading1: {
                        run: {
                            size: 48, // 24pt
                            font: 'Segoe UI Light',
                            color: '0078D4',
                            bold: false,
                        },
                        paragraph: {
                            spacing: {
                                before: 480,
                                after: 240,
                            },
                        },
                    },
                    heading2: {
                        run: {
                            size: 32, // 16pt
                            font: 'Segoe UI',
                            color: '323130',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 360,
                                after: 180,
                            },
                        },
                    },
                },
                table: {
                    borders: {
                        top: { style: docx_1.BorderStyle.NONE },
                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '0078D4' },
                        left: { style: docx_1.BorderStyle.NONE },
                        right: { style: docx_1.BorderStyle.NONE },
                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'EDEBE9' },
                        insideVertical: { style: docx_1.BorderStyle.NONE },
                    },
                },
            },
            pageSettings: {
                margins: {
                    top: 1440,
                    right: 1440,
                    bottom: 1440,
                    left: 1440,
                },
                orientation: 'portrait',
            },
            sections: {
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            right: 1440,
                            bottom: 1440,
                            left: 1440,
                        },
                    },
                },
            },
        };
    }
    /**
     * Classic template
     */
    static createClassicTemplate() {
        return {
            styles: {
                default: {
                    document: {
                        run: {
                            size: 24, // 12pt
                            font: 'Book Antiqua',
                            color: '000000',
                        },
                        paragraph: {
                            spacing: {
                                line: 276,
                                after: 120,
                            },
                        },
                    },
                },
                headings: {
                    heading1: {
                        run: {
                            size: 36, // 18pt
                            font: 'Book Antiqua',
                            color: '800000',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 360,
                                after: 180,
                            },
                            alignment: docx_1.AlignmentType.CENTER,
                        },
                    },
                    heading2: {
                        run: {
                            size: 28, // 14pt
                            font: 'Book Antiqua',
                            color: '800000',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 240,
                                after: 120,
                            },
                        },
                    },
                },
                table: {
                    borders: {
                        top: { style: docx_1.BorderStyle.DOUBLE, size: 3, color: '800000' },
                        bottom: { style: docx_1.BorderStyle.DOUBLE, size: 3, color: '800000' },
                        left: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '800000' },
                        right: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '800000' },
                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
                    },
                },
            },
            pageSettings: {
                margins: {
                    top: 1800,
                    right: 1440,
                    bottom: 1800,
                    left: 1440,
                },
                orientation: 'portrait',
            },
            sections: {
                properties: {
                    page: {
                        margin: {
                            top: 1800,
                            right: 1440,
                            bottom: 1800,
                            left: 1440,
                        },
                    },
                },
            },
        };
    }
    /**
     * Simple template
     */
    static createSimpleTemplate() {
        return {
            styles: {
                default: {
                    document: {
                        run: {
                            size: 24, // 12pt
                            font: 'Arial',
                            color: '000000',
                        },
                        paragraph: {
                            spacing: {
                                line: 276,
                                after: 120,
                            },
                        },
                    },
                },
                headings: {
                    heading1: {
                        run: {
                            size: 32, // 16pt
                            font: 'Arial',
                            color: '000000',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 240,
                                after: 120,
                            },
                        },
                    },
                    heading2: {
                        run: {
                            size: 28, // 14pt
                            font: 'Arial',
                            color: '000000',
                            bold: true,
                        },
                        paragraph: {
                            spacing: {
                                before: 180,
                                after: 60,
                            },
                        },
                    },
                },
                table: {
                    borders: {
                        top: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '000000' },
                        bottom: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '000000' },
                        left: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '000000' },
                        right: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '000000' },
                        insideHorizontal: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '000000' },
                        insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1, color: '000000' },
                    },
                },
            },
            pageSettings: {
                margins: {
                    top: 1440,
                    right: 1440,
                    bottom: 1440,
                    left: 1440,
                },
                orientation: 'portrait',
            },
            sections: {
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            right: 1440,
                            bottom: 1440,
                            left: 1440,
                        },
                    },
                },
            },
        };
    }
}
exports.DocumentTemplates = DocumentTemplates;
//# sourceMappingURL=document-templates.js.map