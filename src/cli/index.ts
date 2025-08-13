#!/usr/bin/env node

import { Command } from 'commander';
import * as fs from 'fs-extra';
import * as path from 'path';
import chalk from 'chalk';
import ora from 'ora';
import { MarkdownDocxConverter } from '../index';
import { MarkdownToDocxOptions, DocxToMarkdownOptions, DocumentTemplate, MermaidTheme } from '../types';
import { FileUtils, PerformanceUtils } from '../utils';
import { logger } from '../utils/logger';

const program = new Command();

/**
 * CLI for Markdown-DOCX converter
 */
class ConverterCLI {
  private converter: MarkdownDocxConverter;

  constructor() {
    this.converter = new MarkdownDocxConverter();
  }

  /**
   * Initialize CLI
   */
  init(): void {
    program
      .name('md-docx')
      .description('Markdown <-> DOCX converter with Mermaid diagrams and advanced styling')
      .version('1.0.0');

    this.setupCommands();
    program.parse();
  }

  /**
   * Setup CLI commands
   */
  private setupCommands(): void {
    // Convert Markdown to DOCX
    program
      .command('convert')
      .description('Convert Markdown file to DOCX')
      .argument('<input>', 'Input Markdown file')
      .option('-o, --output <file>', 'Output DOCX file')
  .option('-t, --template <template>', 'Document template', 'modern')
      .option('-m, --mermaid-theme <theme>', 'Mermaid diagram theme', 'default')
      .option('--title <title>', 'Document title')
      .option('--author <author>', 'Document author')
      .option('--subject <subject>', 'Document subject')
      .option('--toc', 'Generate table of contents')
      .option('--no-links', 'Disable link processing')
      .option('--orientation <orientation>', 'Page orientation (portrait|landscape)', 'portrait')
      .option('--verbose', 'Verbose output')
      .action(async (input, options) => {
        await this.handleConvertCommand(input, options);
      });

    // Extract DOCX to Markdown
    program
      .command('extract')
      .description('Extract DOCX file to Markdown')
      .argument('<input>', 'Input DOCX file')
      .option('-o, --output <file>', 'Output Markdown file')
      .option('--extract-images', 'Extract images to separate files')
      .option('--image-dir <dir>', 'Directory for extracted images', 'images')
      .option('--image-format <format>', 'Image format (png|jpg|svg)', 'png')
      .option('--preserve-formatting', 'Preserve original formatting')
      .option('--verbose', 'Verbose output')
      .action(async (input, options) => {
        await this.handleExtractCommand(input, options);
      });

    // Batch conversion
    program
      .command('batch')
      .description('Batch convert multiple files')
      .argument('<input-dir>', 'Input directory')
      .option('-o, --output-dir <dir>', 'Output directory', './output')
      .option('-f, --format <format>', 'Target format (docx|markdown)', 'docx')
  .option('-t, --template <template>', 'Document template (for DOCX)', 'modern')
      .option('-m, --mermaid-theme <theme>', 'Mermaid theme (for DOCX)', 'default')
      .option('--toc', 'Generate table of contents (for DOCX)')
      .option('--extract-images', 'Extract images (for Markdown)')
      .option('--verbose', 'Verbose output')
      .action(async (inputDir, options) => {
        await this.handleBatchCommand(inputDir, options);
      });

    // Validate Markdown
    program
      .command('validate')
      .description('Validate Markdown file for conversion')
      .argument('<input>', 'Input Markdown file')
      .option('--verbose', 'Verbose output')
      .action(async (input, options) => {
        await this.handleValidateCommand(input, options);
      });

    // Show file statistics
    program
      .command('stats')
      .description('Show file statistics')
      .argument('<input>', 'Input file')
      .option('--verbose', 'Verbose output')
      .action(async (input, options) => {
        await this.handleStatsCommand(input, options);
      });

    // List available templates and themes
    program
      .command('list')
      .description('List available templates and themes')
      .action(() => {
        this.handleListCommand();
      });
  }

  /**
   * Handle convert command
   */
  private async handleConvertCommand(input: string, options: any): Promise<void> {
    try {
      this.setupLogging(options.verbose);
      
      const spinner = ora('Converting Markdown to DOCX...').start();

      // Validate input file
      const inputPath = path.resolve(input);
      if (!await fs.pathExists(inputPath)) {
        spinner.fail(`Input file not found: ${inputPath}`);
        process.exit(1);
      }

      // Determine output file
      const outputPath = options.output 
        ? path.resolve(options.output)
        : inputPath.replace(/\.(md|markdown)$/i, '.docx');

      // Prepare conversion options
      const conversionOptions: MarkdownToDocxOptions = {
        template: this.validateTemplate(options.template),
        mermaidTheme: this.validateMermaidTheme(options.mermaidTheme),
        title: options.title,
        author: options.author,
        subject: options.subject,
        tocGeneration: options.toc,
        preserveLinks: !options.noLinks,
        orientation: options.orientation === 'landscape' ? 'landscape' : 'portrait',
      };

      // Perform conversion
      const startTime = Date.now();
      const result = await this.converter.markdownFileToDocx(inputPath, outputPath, conversionOptions);

      if (result.success) {
        const elapsed = Date.now() - startTime;
        spinner.succeed(`Conversion completed in ${elapsed}ms`);
        
        console.log(chalk.green('   - Success:'), `DOCX file created: ${chalk.cyan(outputPath)}`);
        
        if (result.metadata) {
          console.log(chalk.gray('Statistics:'));
          console.log(`  Input size: ${PerformanceUtils.formatBytes(result.metadata.inputSize)}`);
          console.log(`  Output size: ${PerformanceUtils.formatBytes(result.metadata.outputSize)}`);
          console.log(`  Processing time: ${result.metadata.processingTime}ms`);
          
          if (result.metadata.mermaidDiagramCount > 0) {
            console.log(`  Mermaid diagrams: ${result.metadata.mermaidDiagramCount}`);
          }
          
          if (result.metadata.internalLinkCount > 0) {
            console.log(`  Internal links: ${result.metadata.internalLinkCount}`);
          }
          
          if (result.metadata.externalLinkCount > 0) {
            console.log(`  External links: ${result.metadata.externalLinkCount}`);
          }
        }

        if (result.warnings && result.warnings.length > 0) {
          console.log(chalk.yellow('Warnings:'));
          result.warnings.forEach(warning => console.log(`     - ${warning}`));
        }
      } else {
        spinner.fail('Conversion failed');
        console.error(chalk.red('   - Error:'), result.error?.message || 'Unknown error');
        process.exit(1);
      }
    } catch (error) {
      console.error(chalk.red('   - Fatal error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  }

  /**
   * Handle extract command
   */
  private async handleExtractCommand(input: string, options: any): Promise<void> {
    try {
      this.setupLogging(options.verbose);
      
      const spinner = ora('Extracting DOCX to Markdown...').start();

      // Validate input file
      const inputPath = path.resolve(input);
      if (!await fs.pathExists(inputPath)) {
        spinner.fail(`Input file not found: ${inputPath}`);
        process.exit(1);
      }

      // Determine output file
      const outputPath = options.output 
        ? path.resolve(options.output)
        : inputPath.replace(/\.docx$/i, '.md');

      // Prepare conversion options
      const conversionOptions: DocxToMarkdownOptions = {
        extractImages: options.extractImages,
        imageOutputDir: options.imageDir,
        imageFormat: options.imageFormat,
        preserveFormatting: options.preserveFormatting,
      };

      // Perform conversion
      const startTime = Date.now();
      const result = await this.converter.docxFileToMarkdown(inputPath, outputPath, conversionOptions);

      if (result.success) {
        const elapsed = Date.now() - startTime;
        spinner.succeed(`Extraction completed in ${elapsed}ms`);
        
        console.log(chalk.green('   - Success:'), `Markdown file created: ${chalk.cyan(outputPath)}`);
        
        if (result.metadata) {
          console.log(chalk.gray('Statistics:'));
          console.log(`  Input size: ${PerformanceUtils.formatBytes(result.metadata.inputSize)}`);
          console.log(`  Output size: ${PerformanceUtils.formatBytes(result.metadata.outputSize)}`);
          console.log(`  Processing time: ${result.metadata.processingTime}ms`);
          
          if (result.metadata.imageCount > 0) {
            console.log(`  Images extracted: ${result.metadata.imageCount}`);
          }
        }

        if (result.warnings && result.warnings.length > 0) {
          console.log(chalk.yellow('Warnings:'));
          result.warnings.forEach(warning => console.log(`     - ${warning}`));
        }
      } else {
        spinner.fail('Extraction failed');
        console.error(chalk.red('   - Error:'), result.error?.message || 'Unknown error');
        process.exit(1);
      }
    } catch (error) {
      console.error(chalk.red('   - Fatal error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  }

  /**
   * Handle batch command
   */
  private async handleBatchCommand(inputDir: string, options: any): Promise<void> {
    try {
      this.setupLogging(options.verbose);
      
      const inputPath = path.resolve(inputDir);
      const outputPath = path.resolve(options.outputDir);

      if (!await fs.pathExists(inputPath)) {
        console.error(chalk.red('   - Error:'), `Input directory not found: ${inputPath}`);
        process.exit(1);
      }

      // Find input files
      const fileExtension = options.format === 'docx' ? ['.md', '.markdown'] : ['.docx'];
      const inputFiles = await this.findFiles(inputPath, fileExtension);

      if (inputFiles.length === 0) {
        console.log(chalk.yellow('No files found for conversion'));
        return;
      }

      console.log(`Found ${inputFiles.length} files for batch conversion`);
      
      const spinner = ora(`Converting ${inputFiles.length} files...`).start();

      let results: Array<{ input: string; output: string; result: any }>;
      
      if (options.format === 'docx') {
        const conversionOptions: MarkdownToDocxOptions = {
          template: this.validateTemplate(options.template),
          mermaidTheme: this.validateMermaidTheme(options.mermaidTheme),
          tocGeneration: options.toc,
        };
        
        results = await this.converter.batchMarkdownToDocx(inputFiles, outputPath, conversionOptions);
      } else {
        const conversionOptions: DocxToMarkdownOptions = {
          extractImages: options.extractImages,
        };
        
        results = await this.converter.batchDocxToMarkdown(inputFiles, outputPath, conversionOptions);
      }

      const successful = results.filter(r => r.result.success).length;
      const failed = results.length - successful;

      if (failed === 0) {
        spinner.succeed(`Batch conversion completed: ${successful} files converted`);
      } else {
        spinner.warn(`Batch conversion completed: ${successful} success, ${failed} failed`);
      }

      // Show results summary
      console.log(chalk.green(`   - Successful: ${successful}`));
      if (failed > 0) {
        console.log(chalk.red(`   - Failed: ${failed}`));
        
        const failedFiles = results.filter(r => !r.result.success);
        failedFiles.forEach(({ input, result }) => {
          console.log(`  ${chalk.red('FAILED')} ${input}: ${result.error?.message || 'Unknown error'}`);
        });
      }

    } catch (error) {
      console.error(chalk.red('   - Fatal error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  }

  /**
   * Handle validate command
   */
  private async handleValidateCommand(input: string, options: any): Promise<void> {
    try {
      this.setupLogging(options.verbose);
      
      const inputPath = path.resolve(input);
      if (!await fs.pathExists(inputPath)) {
        console.error(chalk.red('   - Error:'), `Input file not found: ${inputPath}`);
        process.exit(1);
      }

      const spinner = ora('Validating Markdown file...').start();

      const content = await FileUtils.readFile(inputPath);
      const validation = await this.converter.validateMarkdown(content);

      if (validation.isValid) {
        spinner.succeed('Markdown validation passed');
        console.log(chalk.green('   - Valid:'), 'Markdown file is ready for conversion');
      } else {
        spinner.warn('Markdown validation completed with warnings');
        console.log(chalk.yellow('   - Warnings:'));
        validation.warnings.forEach(warning => console.log(`     - ${warning}`));
      }

      if (validation.suggestions.length > 0) {
        console.log(chalk.blue('    - Suggestions:'));
        validation.suggestions.forEach(suggestion => console.log(`     - ${suggestion}`));
      }

    } catch (error) {
      console.error(chalk.red('   - Fatal error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  }

  /**
   * Handle stats command
   */
  private async handleStatsCommand(input: string, options: any): Promise<void> {
    try {
      this.setupLogging(options.verbose);
      
      const inputPath = path.resolve(input);
      if (!await fs.pathExists(inputPath)) {
        console.error(chalk.red('   - Error:'), `Input file not found: ${inputPath}`);
        process.exit(1);
      }

      const spinner = ora('Analyzing file...').start();

      const stats = await this.converter.getConversionStats(inputPath);
      spinner.succeed('File analysis completed');

      console.log(chalk.cyan('    - File Statistics:'));
      console.log(`  File size: ${stats.fileSize}`);
      console.log(`  Word count: ${stats.wordCount.toLocaleString()}`);
      console.log(`  Reading time: ${stats.readingTime} minute(s)`);
      console.log(`  Line count: ${stats.lineCount.toLocaleString()}`);
      console.log(`  Images: ${stats.imageCount}`);
      console.log(`  Links: ${stats.linkCount}`);

    } catch (error) {
      console.error(chalk.red('   - Fatal error:'), error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
  }

  /**
   * Handle list command
   */
  private handleListCommand(): void {
    console.log(chalk.cyan('    - Available Templates:'));
    const templates = this.converter.getAvailableTemplates();
    templates.forEach(template => {
      console.log(`     - ${template}`);
    });

    console.log(chalk.cyan('\n    - Available Mermaid Themes:'));
    const themes = this.converter.getAvailableMermaidThemes();
    themes.forEach(theme => {
      console.log(`     - ${theme}`);
    });
  }

  /**
   * Find files in directory
   */
  private async findFiles(dir: string, extensions: string[]): Promise<string[]> {
    const files: string[] = [];
    const items = await fs.readdir(dir);

    for (const item of items) {
      const itemPath = path.join(dir, item);
      const stat = await fs.stat(itemPath);

      if (stat.isDirectory()) {
        const subFiles = await this.findFiles(itemPath, extensions);
        files.push(...subFiles);
      } else if (extensions.some(ext => item.toLowerCase().endsWith(ext))) {
        files.push(itemPath);
      }
    }

    return files;
  }

  /**
   * Validate template name
   */
  private validateTemplate(template: string): DocumentTemplate {
    const validTemplates = this.converter.getAvailableTemplates();
    if (validTemplates.includes(template as DocumentTemplate)) {
      return template as DocumentTemplate;
    }
    
  console.warn(chalk.yellow(`Invalid template '${template}', using 'modern'`));
  return 'modern';
  }

  /**
   * Validate Mermaid theme
   */
  private validateMermaidTheme(theme: string): MermaidTheme {
    const validThemes = this.converter.getAvailableMermaidThemes();
    if (validThemes.includes(theme as MermaidTheme)) {
      return theme as MermaidTheme;
    }
    
    console.warn(chalk.yellow(`Invalid Mermaid theme '${theme}', using 'default'`));
    return 'default';
  }

  /**
   * Setup logging based on verbosity
   */
  private setupLogging(verbose: boolean): void {
    // Logger level setup will be handled internally
    if (verbose) {
      // Enable debug logging
      console.log('Debug logging enabled');
    }
  }
}

// Initialize and run CLI
const cli = new ConverterCLI();
cli.init();
