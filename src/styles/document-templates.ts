import {
  Document,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  UnderlineType,
  BorderStyle,
  WidthType,
  TableRow,
  TableCell,
  Table,
  Header,
  Footer,
  PageNumber,
  NumberFormat,
} from 'docx';
import { DocumentTemplate, CustomStyles, PageMargins } from '../types';

/**
 * Professional document templates for DOCX generation
 */
export class DocumentTemplates {
  /**
   * Get template configuration
   */
  static getTemplate(templateName: DocumentTemplate): DocumentTemplateConfig {
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
  private static createProfessionalReportTemplate(): DocumentTemplateConfig {
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
                  style: BorderStyle.SINGLE,
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
              alignment: AlignmentType.JUSTIFIED,
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
                style: BorderStyle.SINGLE,
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
            top: { style: BorderStyle.SINGLE, size: 1, color: '2F5597' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: '2F5597' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '2F5597' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '2F5597' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          },
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
        },
        hyperlink: {
          run: {
            color: '0563C1',
            underline: {
              type: UnderlineType.SINGLE,
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
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Professional Report",
                    size: 20,
                    color: '666666',
                  }),
                ],
                alignment: AlignmentType.RIGHT,
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Page ",
                    size: 20,
                    color: '666666',
                  }),
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    size: 20,
                    color: '666666',
                  }),
                  new TextRun({
                    text: " of ",
                    size: 20,
                    color: '666666',
                  }),
                  new TextRun({
                    children: [PageNumber.TOTAL_PAGES],
                    size: 20,
                    color: '666666',
                  }),
                ],
                alignment: AlignmentType.CENTER,
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
  private static createTechnicalDocumentationTemplate(): DocumentTemplateConfig {
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
              top: { color: 'E0E0E0', space: 1, style: BorderStyle.SINGLE, size: 4 },
              bottom: { color: 'E0E0E0', space: 1, style: BorderStyle.SINGLE, size: 4 },
              left: { color: '1E88E5', space: 1, style: BorderStyle.SINGLE, size: 12 },
              right: { color: 'E0E0E0', space: 1, style: BorderStyle.SINGLE, size: 4 },
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
            top: { style: BorderStyle.SINGLE, size: 2, color: '1E88E5' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'BDBDBD' },
            left: { style: BorderStyle.SINGLE, size: 1, color: 'BDBDBD' },
            right: { style: BorderStyle.SINGLE, size: 1, color: 'BDBDBD' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'E0E0E0' },
          },
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
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
  private static createBusinessProposalTemplate(): DocumentTemplateConfig {
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
              alignment: AlignmentType.CENTER,
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
            top: { style: BorderStyle.DOUBLE, size: 4, color: '8B4513' },
            bottom: { style: BorderStyle.DOUBLE, size: 4, color: '8B4513' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '8B4513' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '8B4513' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'DDD' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'DDD' },
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
  private static createAcademicPaperTemplate(): DocumentTemplateConfig {
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
              alignment: AlignmentType.CENTER,
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
            top: { style: BorderStyle.SINGLE, size: 2, color: '000000' },
            bottom: { style: BorderStyle.SINGLE, size: 2, color: '000000' },
            left: { style: BorderStyle.NONE },
            right: { style: BorderStyle.NONE },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            insideVertical: { style: BorderStyle.NONE },
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
  private static createModernTemplate(): DocumentTemplateConfig {
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
            top: { style: BorderStyle.NONE },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: '0078D4' },
            left: { style: BorderStyle.NONE },
            right: { style: BorderStyle.NONE },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'EDEBE9' },
            insideVertical: { style: BorderStyle.NONE },
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
  private static createClassicTemplate(): DocumentTemplateConfig {
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
              alignment: AlignmentType.CENTER,
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
            top: { style: BorderStyle.DOUBLE, size: 3, color: '800000' },
            bottom: { style: BorderStyle.DOUBLE, size: 3, color: '800000' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '800000' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '800000' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
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
  private static createSimpleTemplate(): DocumentTemplateConfig {
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
            top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
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
