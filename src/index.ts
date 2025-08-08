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

import { MarkdownToDocxConverter } from './converters/markdown-to-docx';
import { DocxToMarkdownConverter } from './converters/docx-to-markdown';
import {
  MarkdownToDocxOptions,
  DocxToMarkdownOptions,
  ConversionResult,
  DocumentTemplate,
  MermaidTheme,
} from './types';
import { FileUtils, PerformanceUtils } from './utils';
import { logger } from './utils/logger';

/**
 * Main converter class providing simplified API
 */
export class MarkdownDocxConverter {
  private markdownToDocxConverter: MarkdownToDocxConverter | null = null;
  private docxToMarkdownConverter: DocxToMarkdownConverter | null = null;

  constructor(
    private defaultOptions: {
      template?: DocumentTemplate;
      mermaidTheme?: MermaidTheme;
      preserveLinks?: boolean;
      tocGeneration?: boolean;
    } = {}
  ) {}

  /**
   * Convert Markdown content to DOCX buffer
   */
  async markdownToDocx(
    markdownContent: string,
    options: MarkdownToDocxOptions = {}
  ): Promise<ConversionResult> {
    const mergedOptions = { ...this.defaultOptions, ...options };
    
    this.markdownToDocxConverter = new MarkdownToDocxConverter();
    
    try {
      return await this.markdownToDocxConverter.convert(markdownContent, mergedOptions);
    } finally {
      this.markdownToDocxConverter = null;
    }
  }

  /**
   * Convert Markdown file to DOCX file
   */
  async markdownFileToDocx(
    inputFilePath: string,
    outputFilePath: string,
    options: MarkdownToDocxOptions = {}
  ): Promise<ConversionResult> {
    try {
      logger.info('Converting Markdown file to DOCX', { input: inputFilePath, output: outputFilePath });

      // Validate input file
      const validation = FileUtils.validateFile(inputFilePath, ['.md', '.markdown']);
      if (!validation.isValid) {
        throw new Error(`Invalid input file: ${validation.errors.map(e => e.message).join(', ')}`);
      }

      // Read markdown content
      const markdownContent = await FileUtils.readFile(inputFilePath);

      // Convert to DOCX
      const result = await this.markdownToDocx(markdownContent, options);

      if (result.success && result.output) {
        // Save DOCX file
        await FileUtils.writeFile(outputFilePath, result.output);
        logger.info('DOCX file saved successfully', { path: outputFilePath });
      }

      return result;
    } catch (error) {
      logger.error('Failed to convert Markdown file to DOCX', { error, input: inputFilePath });
      
      return {
        success: false,
        error: {
          code: 'FILE_CONVERSION_FAILED',
          message: error instanceof Error ? error.message : String(error),
          details: { input: inputFilePath, output: outputFilePath },
        },
      };
    }
  }

  /**
   * Convert DOCX file to Markdown content
   */
  async docxToMarkdown(
    docxFilePath: string,
    options: DocxToMarkdownOptions = {}
  ): Promise<ConversionResult> {
    const mergedOptions = { ...this.defaultOptions, ...options };
    
    this.docxToMarkdownConverter = new DocxToMarkdownConverter(mergedOptions);
    
    try {
      return await this.docxToMarkdownConverter.convert(docxFilePath, mergedOptions);
    } finally {
      this.docxToMarkdownConverter = null;
    }
  }

  /**
   * Convert DOCX file to Markdown file
   */
  async docxFileToMarkdown(
    inputFilePath: string,
    outputFilePath: string,
    options: DocxToMarkdownOptions = {}
  ): Promise<ConversionResult> {
    try {
      logger.info('Converting DOCX file to Markdown', { input: inputFilePath, output: outputFilePath });

      // Convert to Markdown
      const result = await this.docxToMarkdown(inputFilePath, options);

      if (result.success && result.output) {
        // Save Markdown file
        await FileUtils.writeFile(outputFilePath, result.output as string);
        logger.info('Markdown file saved successfully', { path: outputFilePath });
      }

      return result;
    } catch (error) {
      logger.error('Failed to convert DOCX file to Markdown', { error, input: inputFilePath });
      
      return {
        success: false,
        error: {
          code: 'FILE_CONVERSION_FAILED',
          message: error instanceof Error ? error.message : String(error),
          details: { input: inputFilePath, output: outputFilePath },
        },
      };
    }
  }

  /**
   * Batch convert multiple Markdown files to DOCX
   */
  async batchMarkdownToDocx(
    inputFiles: string[],
    outputDir: string,
    options: MarkdownToDocxOptions = {}
  ): Promise<Array<{ input: string; output: string; result: ConversionResult }>> {
    logger.info(`Starting batch conversion of ${inputFiles.length} Markdown files`);

    await FileUtils.ensureWritableDir(outputDir);

    const results: Array<{ input: string; output: string; result: ConversionResult }> = [];

    for (const inputFile of inputFiles) {
      const fileName = FileUtils.getRelativePath(inputFile, inputFile).replace(/\.(md|markdown)$/i, '.docx');
      const outputFile = `${outputDir}/${fileName}`;

      try {
        const result = await this.markdownFileToDocx(inputFile, outputFile, options);
        results.push({ input: inputFile, output: outputFile, result });
      } catch (error) {
        logger.error(`Failed to convert file: ${inputFile}`, { error });
        results.push({
          input: inputFile,
          output: outputFile,
          result: {
            success: false,
            error: {
              code: 'BATCH_CONVERSION_ERROR',
              message: error instanceof Error ? error.message : String(error),
            },
          },
        });
      }
    }

    const successful = results.filter(r => r.result.success).length;
    logger.info(`Batch conversion completed: ${successful}/${inputFiles.length} files converted successfully`);

    return results;
  }

  /**
   * Batch convert multiple DOCX files to Markdown
   */
  async batchDocxToMarkdown(
    inputFiles: string[],
    outputDir: string,
    options: DocxToMarkdownOptions = {}
  ): Promise<Array<{ input: string; output: string; result: ConversionResult }>> {
    logger.info(`Starting batch conversion of ${inputFiles.length} DOCX files`);

    await FileUtils.ensureWritableDir(outputDir);

    const results: Array<{ input: string; output: string; result: ConversionResult }> = [];

    for (const inputFile of inputFiles) {
      const fileName = FileUtils.getRelativePath(inputFile, inputFile).replace(/\.docx$/i, '.md');
      const outputFile = `${outputDir}/${fileName}`;

      try {
        const result = await this.docxFileToMarkdown(inputFile, outputFile, options);
        results.push({ input: inputFile, output: outputFile, result });
      } catch (error) {
        logger.error(`Failed to convert file: ${inputFile}`, { error });
        results.push({
          input: inputFile,
          output: outputFile,
          result: {
            success: false,
            error: {
              code: 'BATCH_CONVERSION_ERROR',
              message: error instanceof Error ? error.message : String(error),
            },
          },
        });
      }
    }

    const successful = results.filter(r => r.result.success).length;
    logger.info(`Batch conversion completed: ${successful}/${inputFiles.length} files converted successfully`);

    return results;
  }

  /**
   * Get conversion statistics
   */
  async getConversionStats(filePath: string): Promise<{
    fileSize: string;
    wordCount: number;
    readingTime: number;
    lineCount: number;
    imageCount: number;
    linkCount: number;
  }> {
    try {
      const content = await FileUtils.readFile(filePath);
      const stats = await PerformanceUtils.measureAsync(async () => {
        const wordCount = content.split(/\s+/).length;
        const readingTime = Math.ceil(wordCount / 200); // 200 WPM
        const lineCount = content.split('\n').length;
        const imageCount = (content.match(/!\[.*?\]\(.*?\)/g) || []).length;
        // Count standard Markdown links [text](url) excluding images ![alt](src)
        let linkCount = 0;
        const linkRe = /\[[^\]]*\]\([^)]*\)/g;
        for (const m of content.matchAll(linkRe) as any) {
          const idx = m.index ?? 0;
          const isImage = idx > 0 && content.charAt(idx - 1) === '!';
          if (!isImage) linkCount += 1;
        }
        
        return {
          wordCount,
          readingTime,
          lineCount,
          imageCount,
          linkCount,
        };
      });

      const fileSize = PerformanceUtils.formatBytes(Buffer.from(content, 'utf-8').length);

      return {
        fileSize,
        ...stats.result,
      };
    } catch (error) {
      logger.error('Failed to get conversion stats', { error, file: filePath });
      throw error;
    }
  }

  /**
   * Validate Markdown content for conversion
   */
  async validateMarkdown(content: string): Promise<{
    isValid: boolean;
    warnings: string[];
    suggestions: string[];
  }> {
    const warnings: string[] = [];
    const suggestions: string[] = [];

    try {
      // Check for common issues
      const lines = content.split('\n');
      
      // Check for proper heading structure
      const headings = lines.filter(line => line.match(/^#{1,6}\s/));
      if (headings.length === 0) {
        warnings.push('No headings found - consider adding section headers');
      }

      // Check for broken links
      const links = content.match(/\[.*?\]\(.*?\)/g) || [];
      for (const link of links) {
        const urlMatch = link.match(/\((.*?)\)/);
        if (urlMatch && urlMatch[1]) {
          const url = urlMatch[1];
          if (url.startsWith('http') && !url.includes('://')) {
            warnings.push(`Potentially malformed URL: ${url}`);
          }
        }
      }

      // Check for Mermaid diagrams
      const mermaidBlocks = content.match(/```mermaid[\s\S]*?```/g) || [];
      if (mermaidBlocks.length > 0) {
        suggestions.push(`Found ${mermaidBlocks.length} Mermaid diagram(s) - will be converted to images`);
      }

      // Check for tables
      const tables = content.match(/\|.*\|/g) || [];
      if (tables.length > 0) {
        suggestions.push(`Found ${tables.length} table row(s) - formatting will be preserved`);
      }

      return {
        isValid: warnings.length === 0,
        warnings,
        suggestions,
      };
    } catch (error) {
      logger.error('Failed to validate Markdown', { error });
      
      return {
        isValid: false,
        warnings: ['Failed to validate Markdown content'],
        suggestions: [],
      };
    }
  }

  /**
   * Get available templates
   */
  getAvailableTemplates(): DocumentTemplate[] {
    return [
      'simple',
      'professional-report',
      'technical-documentation',
      'business-proposal',
      'academic-paper',
      'modern',
      'classic',
    ];
  }

  /**
   * Get available Mermaid themes
   */
  getAvailableMermaidThemes(): MermaidTheme[] {
    return ['default', 'forest', 'dark', 'neutral', 'base'];
  }
}

// Default export for convenience
export default MarkdownDocxConverter;
