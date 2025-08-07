import { ProcessedLink } from '../types';
import { StringUtils } from '../utils';
import { logger } from '../utils/logger';

/**
 * Advanced link processor for Markdown and DOCX conversion
 */
export class LinkProcessor {
  private anchors: Map<string, string> = new Map();
  private externalLinks: ProcessedLink[] = [];
  private internalLinks: ProcessedLink[] = [];

  /**
   * Process content containing links
   */
  processContent(content: string): { content: string; links: any[]; headings: any[] } {
    // Simple implementation - just extract basic link info
    const linkMatches = content.match(/\[([^\]]+)\]\(([^)]+)\)/g) || [];
    const headingMatches = content.match(/^#{1,6}\s+(.+)$/gm) || [];
    
    const links = linkMatches.map(match => {
      const [, text, url] = match.match(/\[([^\]]+)\]\(([^)]+)\)/) || [];
      return {
        text,
        url,
        type: url?.startsWith('http') ? 'external' : 'internal'
      };
    });

    const headings = headingMatches.map(match => {
      const [, text] = match.match(/^#{1,6}\s+(.+)$/) || [];
      return { text };
    });

    return { content, links, headings };
  }

  /**
   * Process all links in Markdown content
   */
  processMarkdownLinks(markdown: string): {
    processedContent: string;
    links: ProcessedLink[];
  } {
    let processedContent = markdown;
    const allLinks: ProcessedLink[] = [];

    // Extract headings to build anchor map
    this.buildAnchorMap(markdown);

    // Process different types of links
    processedContent = this.processInternalLinks(processedContent, allLinks);
    processedContent = this.processExternalLinks(processedContent, allLinks);
    processedContent = this.processReferenceLinks(processedContent, allLinks);
    processedContent = this.processImageLinks(processedContent, allLinks);

    logger.info(`Processed ${allLinks.length} links`, {
      internal: allLinks.filter(l => l.type === 'internal').length,
      external: allLinks.filter(l => l.type === 'external').length,
      anchors: allLinks.filter(l => l.type === 'anchor').length,
    });

    return { processedContent, links: allLinks };
  }

  /**
   * Convert Markdown links to Word hyperlinks
   */
  convertToWordLinks(links: ProcessedLink[]): Array<{
    text: string;
    url: string;
    type: 'hyperlink' | 'bookmark';
    style?: string;
  }> {
    return links.map(link => {
      if (link.type === 'external') {
        return {
          text: link.text,
          url: link.url,
          type: 'hyperlink' as const,
          style: 'Hyperlink',
        };
      } else if (link.type === 'internal' || link.type === 'anchor') {
        // Convert internal links to bookmarks
        const bookmarkName = this.createBookmarkName(link.url);
        return {
          text: link.text,
          url: bookmarkName,
          type: 'bookmark' as const,
          style: 'IntenseReference',
        };
      }

      return {
        text: link.text,
        url: link.url,
        type: 'hyperlink' as const,
      };
    });
  }

  /**
   * Extract links from DOCX content
   */
  extractDocxLinks(content: string): ProcessedLink[] {
    const links: ProcessedLink[] = [];
    
    // This would be implemented with the actual DOCX parsing
    // For now, return empty array as placeholder
    logger.debug('Extracting links from DOCX content');
    
    return links;
  }

  /**
   * Validate link URLs
   */
  validateLinks(links: ProcessedLink[]): ProcessedLink[] {
    return links.map(link => {
      let isValid = true;
      let processedUrl = link.url;

      switch (link.type) {
        case 'external':
          isValid = this.isValidHttpUrl(link.url);
          break;
        
        case 'internal':
        case 'anchor':
          const anchorName = link.url.startsWith('#') ? link.url.substring(1) : link.url;
          isValid = this.anchors.has(anchorName);
          processedUrl = this.anchors.get(anchorName) || link.url;
          break;
      }

      return {
        ...link,
        isValid,
        processedUrl,
      };
    });
  }

  /**
   * Generate anchor map from headings
   */
  private buildAnchorMap(markdown: string): void {
    this.anchors.clear();
    
    const headingRegex = /^(#{1,6})\s+(.+)$/gm;
    let match;

    while ((match = headingRegex.exec(markdown)) !== null) {
      const headingText = match[2].trim();
      const anchor = StringUtils.createAnchor(headingText);
      this.anchors.set(anchor, headingText);
      
      // Also map the original text for direct matches
      this.anchors.set(headingText.toLowerCase().replace(/\s+/g, '-'), headingText);
    }

    logger.debug(`Built anchor map with ${this.anchors.size} entries`);
  }

  /**
   * Process internal links [text](#anchor)
   */
  private processInternalLinks(content: string, links: ProcessedLink[]): string {
    const internalLinkRegex = /\[([^\]]+)\]\(#([^)]+)\)/g;
    
    return content.replace(internalLinkRegex, (match, text, anchor) => {
      const link: ProcessedLink = {
        text: text.trim(),
        url: `#${anchor}`,
        type: 'anchor',
        isValid: this.anchors.has(anchor),
      };

      links.push(link);
      return match; // Keep original format for now
    });
  }

  /**
   * Process external links [text](http://example.com)
   */
  private processExternalLinks(content: string, links: ProcessedLink[]): string {
    const externalLinkRegex = /\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g;
    
    return content.replace(externalLinkRegex, (match, text, url) => {
      const link: ProcessedLink = {
        text: text.trim(),
        url: url.trim(),
        type: 'external',
        isValid: this.isValidHttpUrl(url),
      };

      links.push(link);
      return match; // Keep original format for now
    });
  }

  /**
   * Process reference-style links [text][ref]
   */
  private processReferenceLinks(content: string, links: ProcessedLink[]): string {
    const referenceMap = this.extractReferenceDefinitions(content);
    const referenceLinkRegex = /\[([^\]]+)\]\[([^\]]+)\]/g;
    
    return content.replace(referenceLinkRegex, (match, text, ref) => {
      const refDefinition = referenceMap.get(ref.toLowerCase());
      
      if (refDefinition) {
        const link: ProcessedLink = {
          text: text.trim(),
          url: refDefinition.url,
          type: this.isValidHttpUrl(refDefinition.url) ? 'external' : 'internal',
          isValid: true,
        };

        links.push(link);
        return `[${text}](${refDefinition.url})`;
      }

      return match;
    });
  }

  /**
   * Process image links ![alt](src)
   */
  private processImageLinks(content: string, links: ProcessedLink[]): string {
    const imageLinkRegex = /!\[([^\]]*)\]\(([^)]+)\)/g;
    
    return content.replace(imageLinkRegex, (match, alt, src) => {
      const link: ProcessedLink = {
        text: alt || 'Image',
        url: src.trim(),
        type: this.isValidHttpUrl(src) ? 'external' : 'internal',
        isValid: true, // Images are handled separately
      };

      links.push(link);
      return match; // Keep original format
    });
  }

  /**
   * Extract reference definitions [ref]: url "title"
   */
  private extractReferenceDefinitions(content: string): Map<string, { url: string; title?: string }> {
    const refMap = new Map();
    const refDefRegex = /^\s*\[([^\]]+)\]:\s*([^\s]+)(?:\s+"([^"]*)")?/gm;
    let match;

    while ((match = refDefRegex.exec(content)) !== null) {
      const [, ref, url, title] = match;
      refMap.set(ref.toLowerCase(), {
        url: url.trim(),
        title: title?.trim(),
      });
    }

    return refMap;
  }

  /**
   * Validate HTTP URL
   */
  private isValidHttpUrl(url: string): boolean {
    try {
      const urlObj = new URL(url);
      return urlObj.protocol === 'http:' || urlObj.protocol === 'https:';
    } catch {
      return false;
    }
  }

  /**
   * Create bookmark name for Word document
   */
  private createBookmarkName(anchor: string): string {
    return anchor
      .replace(/^#+/, '')
      .replace(/[^a-zA-Z0-9]/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '')
      .substring(0, 40); // Word bookmark name limit
  }

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
  } {
    const entries: Array<{
      title: string;
      level: number;
      anchor: string;
    }> = [];

    const headingRegex = /^(#{1,6})\s+(.+)$/gm;
    let match;

    while ((match = headingRegex.exec(markdown)) !== null) {
      const level = match[1].length;
      const title = match[2].trim();
      const anchor = StringUtils.createAnchor(title);

      entries.push({
        title,
        level,
        anchor,
      });
    }

    // Generate TOC markdown
    const tocLines = ['# Table of Contents\n'];
    
    for (const entry of entries) {
      const indent = '  '.repeat(entry.level - 1);
      const tocLine = `${indent}- [${entry.title}](#${entry.anchor})`;
      tocLines.push(tocLine);
    }

    return {
      tocMarkdown: tocLines.join('\n'),
      entries,
    };
  }

  /**
   * Update internal links after heading changes
   */
  updateInternalLinks(content: string, headingMap: Map<string, string>): string {
    let updatedContent = content;

    headingMap.forEach((newAnchor, oldAnchor) => {
      const linkRegex = new RegExp(`\\[([^\\]]+)\\]\\(#${oldAnchor}\\)`, 'g');
      updatedContent = updatedContent.replace(linkRegex, `[$1](#${newAnchor})`);
    });

    return updatedContent;
  }

  /**
   * Get link statistics
   */
  getLinkStatistics(links: ProcessedLink[]): {
    total: number;
    byType: Record<string, number>;
    valid: number;
    invalid: number;
  } {
    const stats = {
      total: links.length,
      byType: {} as Record<string, number>,
      valid: 0,
      invalid: 0,
    };

    for (const link of links) {
      stats.byType[link.type] = (stats.byType[link.type] || 0) + 1;
      
      if (link.isValid) {
        stats.valid++;
      } else {
        stats.invalid++;
      }
    }

    return stats;
  }

  /**
   * Clean up resources
   */
  cleanup(): void {
    this.anchors.clear();
    this.externalLinks = [];
    this.internalLinks = [];
  }
}
