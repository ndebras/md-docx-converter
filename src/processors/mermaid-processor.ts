import puppeteer, { Browser, Page } from 'puppeteer';
import { ProcessedMermaidDiagram, MermaidTheme } from '../types';
import { logger } from '../utils/logger';
import { StringUtils, PerformanceUtils } from '../utils';
import { v4 as uuidv4 } from 'uuid';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Advanced Mermaid diagram processor with multiple output formats
 */
export class MermaidProcessor {
  private browser: Browser | null = null;
  private page: Page | null = null;
  private isInitialized = false;

  constructor(private theme: MermaidTheme = 'default') {}

  /**
   * Process content containing Mermaid diagrams
   */
  async processContent(content: string, options: { theme?: MermaidTheme } = {}): Promise<{ content: string; diagramCount: number }> {
    const mermaidBlocks = content.match(/```mermaid[\s\S]*?```/g) || [];
    
    if (mermaidBlocks.length === 0) {
      return { content, diagramCount: 0 };
    }

    let processedContent = content;
    let diagramCount = 0;

    // Process each Mermaid block
    for (const block of mermaidBlocks) {
      try {
        // Extract diagram code
        const diagramCode = block.replace(/```mermaid\s*\n/, '').replace(/\n```$/, '').trim();
        
        if (diagramCode) {
          // Create a better text-based representation for DOCX compatibility
          const textRepresentation = this.createEnhancedTextualDiagramRepresentation(diagramCode);
          processedContent = processedContent.replace(block, textRepresentation);
          diagramCount++;
        }
      } catch (error) {
        console.warn('Failed to process Mermaid diagram:', error);
        // Keep original block if processing fails
      }
    }

    return {
      content: processedContent,
      diagramCount
    };
  }

  /**
   * Create an enhanced textual representation of Mermaid diagrams for DOCX
   */
  private createEnhancedTextualDiagramRepresentation(diagramCode: string): string {
    const diagramType = this.detectDiagramType(diagramCode);
    
    // Extract key elements from the diagram code with improved parsing
    if (diagramCode.includes('graph') || diagramCode.includes('flowchart')) {
      return this.createEnhancedFlowchartText(diagramCode);
    } else if (diagramCode.includes('sequenceDiagram')) {
      return this.createEnhancedSequenceText(diagramCode);
    } else if (diagramCode.includes('classDiagram')) {
      return this.createEnhancedClassDiagramText(diagramCode);
    } else {
      return this.createEnhancedGenericDiagramText(diagramType, diagramCode);
    }
  }

  /**
   * Create enhanced flowchart text representation with better parsing
   */
  private createEnhancedFlowchartText(code: string): string {
    const lines = code.split('\n').map(line => line.trim()).filter(line => line);
    const nodes = new Map<string, string>();
    const connections: string[] = [];
    
    // Extract all nodes from the entire code
    const allNodeMatches = code.match(/([A-Za-z0-9]+)\[(.*?)\]/g);
    if (allNodeMatches) {
      allNodeMatches.forEach(match => {
        const nodeMatch = match.match(/([A-Za-z0-9]+)\[(.*?)\]/);
        if (nodeMatch) {
          nodes.set(nodeMatch[1], nodeMatch[2]);
        }
      });
    }
    
    // Extract connections
    lines.forEach(line => {
      if (line.includes('-->')) {
        const parts = line.split('-->');
        if (parts.length === 2) {
          const fromPart = parts[0].trim();
          const toPart = parts[1].trim();
          
          // Extract node IDs (before the [label] part)
          const fromId = fromPart.match(/([A-Za-z0-9]+)/)?.[1] || fromPart;
          const toId = toPart.match(/([A-Za-z0-9]+)/)?.[1] || toPart;
          
          const fromLabel = nodes.get(fromId) || fromId;
          const toLabel = nodes.get(toId) || toId;
          
          connections.push(`${fromLabel} ??? ${toLabel}`);
        }
      }
    });

    return `

???? **DIAGRAMME DE FLUX**

???? **??l??ments :**
${Array.from(nodes.values()).map(label => `   ??? ${label}`).join('\n')}

???? **Flux :**
${connections.length > 0 ? connections.map(conn => `   ${conn}`).join('\n') : '   Connexions d??finies dans le sch??ma'}

`;
  }

  /**
   * Create enhanced sequence diagram text representation
   */
  private createEnhancedSequenceText(code: string): string {
    const lines = code.split('\n').map(line => line.trim()).filter(line => line);
    const participants = new Set<string>();
    const interactions: string[] = [];
    
    lines.forEach(line => {
      if (line.includes('participant ')) {
        const participant = line.replace('participant ', '').trim();
        participants.add(participant);
      } else if (line.includes('->>') || line.includes('-->>')) {
        const arrow = line.includes('->>') ? '???' : '???';
        const parts = line.split(/->>|-->>/, 2);
        if (parts.length === 2) {
          const from = parts[0].trim();
          const to = parts[1].split(':')[0].trim();
          const message = parts[1].includes(':') ? parts[1].split(':')[1].trim() : '';
          participants.add(from);
          participants.add(to);
          interactions.push(`${from} ${arrow} ${to}${message ? ': ' + message : ''}`);
        }
      }
    });

    return `

???? **DIAGRAMME DE S??QUENCE**

???? **Acteurs :**
${Array.from(participants).map(p => `   ??? ${p}`).join('\n')}

???? **Interactions :**
${interactions.map(inter => `   ${inter}`).join('\n')}

`;
  }

  /**
   * Create enhanced class diagram text representation
   */
  private createEnhancedClassDiagramText(code: string): string {
    const lines = code.split('\n').map(line => line.trim()).filter(line => line);
    const classes = new Set<string>();
    const relationships: string[] = [];
    
    lines.forEach(line => {
      if (line.includes('<|--') || line.includes('*--') || line.includes('o--') || line.includes('-->')) {
        const parts = line.split(/(<\|--|--|\*--|o--|-->)/);
        if (parts.length >= 3) {
          const class1 = parts[0].trim();
          const relationship = parts[1];
          const class2 = parts[2].trim().split(':')[0].trim();
          classes.add(class1);
          classes.add(class2);
          
          let relType = '';
          switch(relationship) {
            case '<|--': relType = 'h??rite de'; break;
            case '*--': relType = 'compose'; break;
            case 'o--': relType = 'agr??ge'; break;
            case '-->': relType = 'utilise'; break;
            default: relType = 'li?? ??'; break;
          }
          
          relationships.push(`${class1} ${relType} ${class2}`);
        }
      }
    });

    return `

???? **DIAGRAMME DE CLASSES**

??????? **Classes :**
${Array.from(classes).map(cls => `   ??? ${cls}`).join('\n')}

???? **Relations :**
${relationships.map(rel => `   ${rel}`).join('\n')}

`;
  }

  /**
   * Create enhanced generic diagram text representation
   */
  private createEnhancedGenericDiagramText(type: string, code: string): string {
    const lineCount = code.split('\n').filter(line => line.trim()).length;
    return `

???? **${type.toUpperCase()}**

???? **Informations :**
   ??? Type : ${type}
   ??? Lignes de code : ${lineCount}
   ??? G??n??rateur : Mermaid

??????  *Diagramme converti automatiquement pour compatibilit?? DOCX*

`;
  }

  /**
   * Create a textual representation of Mermaid diagrams for DOCX
   */
  private createTextualDiagramRepresentation(diagramCode: string): string {
    const diagramType = this.detectDiagramType(diagramCode);
    
    // Extract key elements from the diagram code
    if (diagramCode.includes('graph') || diagramCode.includes('flowchart')) {
      return this.createFlowchartText(diagramCode);
    } else if (diagramCode.includes('sequenceDiagram')) {
      return this.createSequenceText(diagramCode);
    } else if (diagramCode.includes('classDiagram')) {
      return this.createClassDiagramText(diagramCode);
    } else {
      return this.createGenericDiagramText(diagramType, diagramCode);
    }
  }

  /**
   * Create flowchart text representation
   */
  private createFlowchartText(code: string): string {
    // Extract nodes and connections with better parsing
    const lines = code.split('\n').map(line => line.trim()).filter(line => line);
    const nodes = new Map<string, string>(); // id -> label
    const connections: string[] = [];
    
    lines.forEach(line => {
      if (line.includes('-->')) {
        const parts = line.split('-->');
        if (parts.length === 2) {
          const from = this.extractNodeInfo(parts[0].trim());
          const to = this.extractNodeInfo(parts[1].trim());
          
          nodes.set(from.id, from.label);
          nodes.set(to.id, to.label);
          connections.push(`${from.label} ??? ${to.label}`);
        }
      }
    });

    // If no connections found, try to extract nodes from individual lines
    if (nodes.size === 0) {
      lines.forEach(line => {
        const nodeMatch = line.match(/([A-Za-z0-9]+)\[(.*?)\]/);
        if (nodeMatch) {
          nodes.set(nodeMatch[1], nodeMatch[2]);
        }
      });
    }

    return `

???? **DIAGRAMME DE FLUX MERMAID**

???? **??l??ments du processus :**
${Array.from(nodes.values()).map(label => `   ??? ${label}`).join('\n')}

???? **Flux de donn??es :**
${connections.length > 0 ? connections.map(conn => `   ${conn}`).join('\n') : '   Flux d??fini dans le diagramme'}

`;
  }

  /**
   * Extract node information (ID and label)
   */
  private extractNodeInfo(nodeText: string): { id: string, label: string } {
    // Handle patterns like: A[Label], B{Decision}, C((Circle)), etc.
    const patterns = [
      /([A-Za-z0-9]+)\[(.*?)\]/, // A[Label]
      /([A-Za-z0-9]+)\{(.*?)\}/, // B{Decision}
      /([A-Za-z0-9]+)\(\((.*?)\)\)/, // C((Circle))
      /([A-Za-z0-9]+)\((.*?)\)/, // D(Rounded)
      /([A-Za-z0-9]+)>>(.*?)>>/, // E>>Flag>>
    ];
    
    for (const pattern of patterns) {
      const match = nodeText.match(pattern);
      if (match) {
        return { id: match[1], label: match[2] };
      }
    }
    
    // Fallback: treat the whole text as both ID and label
    const cleanText = nodeText.replace(/[^\w\s]/g, '').trim();
    return { id: cleanText, label: cleanText || nodeText };
  }

  /**
   * Create sequence diagram text representation
   */
  private createSequenceText(code: string): string {
    const lines = code.split('\n').map(line => line.trim()).filter(line => line);
    const participants = new Set<string>();
    const interactions: string[] = [];
    
    lines.forEach(line => {
      if (line.includes('participant ')) {
        const participant = line.replace('participant ', '').trim();
        participants.add(participant);
      } else if (line.includes('->>') || line.includes('-->>')) {
        const arrow = line.includes('->>') ? '->' : '-->';
        const parts = line.split(/->>|-->>/, 2);
        if (parts.length === 2) {
          const from = parts[0].trim();
          const to = parts[1].split(':')[0].trim();
          const message = parts[1].includes(':') ? parts[1].split(':')[1].trim() : '';
          participants.add(from);
          participants.add(to);
          interactions.push(`${from} ${arrow} ${to}: ${message}`);
        }
      }
    });

    return `

???? **DIAGRAMME DE S??QUENCE MERMAID**

???? **Participants :**
${Array.from(participants).map(p => `   ??? ${p}`).join('\n')}

???? **Interactions :**
${interactions.map(inter => `   ${inter}`).join('\n')}

`;
  }

  /**
   * Create class diagram text representation
   */
  private createClassDiagramText(code: string): string {
    const lines = code.split('\n').map(line => line.trim()).filter(line => line);
    const classes = new Set<string>();
    const relationships: string[] = [];
    
    lines.forEach(line => {
      if (line.includes('<|--') || line.includes('*--') || line.includes('o--')) {
        const parts = line.split(/(<\|--|--|\*--|o--)/);
        if (parts.length >= 3) {
          const class1 = parts[0].trim();
          const relationship = parts[1];
          const class2 = parts[2].trim();
          classes.add(class1);
          classes.add(class2);
          relationships.push(`${class1} ${relationship} ${class2}`);
        }
      }
    });

    return `

???? **DIAGRAMME DE CLASSES MERMAID**

???? **Classes identifi??es :**
${Array.from(classes).map(cls => `   ??? ${cls}`).join('\n')}

???? **Relations :**
${relationships.map(rel => `   ${rel}`).join('\n')}

`;
  }

  /**
   * Create generic diagram text representation
   */
  private createGenericDiagramText(type: string, code: string): string {
    const lineCount = code.split('\n').length;
    return `

???? **${type.toUpperCase()} MERMAID**

???? **Type :** ${type}
???? **Complexit?? :** ${lineCount} lignes de code
???? **Contenu :** Diagramme g??n??r?? automatiquement

??????  *Ce diagramme a ??t?? d??tect?? et trait?? par le convertisseur Mermaid*

`;
  }

  /**
   * Generate a simple SVG placeholder for Mermaid diagrams
   */
  private generateMermaidSVGPlaceholder(diagramCode: string): string {
    // Create a more visual SVG based on diagram type
    const diagramType = this.detectDiagramType(diagramCode);
    let svg = '';

    if (diagramCode.includes('graph') || diagramCode.includes('flowchart')) {
      // Generate a flowchart-like SVG
      svg = this.generateFlowchartSVG(diagramCode);
    } else if (diagramCode.includes('sequenceDiagram')) {
      // Generate a sequence diagram SVG
      svg = this.generateSequenceSVG(diagramCode);
    } else {
      // Default diagram
      svg = this.generateDefaultDiagramSVG(diagramType);
    }
    
    return Buffer.from(svg).toString('base64');
  }

  /**
   * Generate a flowchart-style SVG
   */
  private generateFlowchartSVG(code: string): string {
    return `
      <svg width="500" height="300" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <marker id="arrowhead" markerWidth="10" markerHeight="7" 
                  refX="10" refY="3.5" orient="auto">
            <polygon points="0 0, 10 3.5, 0 7" fill="#333" />
          </marker>
        </defs>
        
        <!-- Background -->
        <rect width="500" height="300" fill="#f8f9fa" stroke="#dee2e6" stroke-width="1"/>
        
        <!-- Nodes -->
        <rect x="50" y="50" width="100" height="40" rx="5" fill="#007bff" stroke="#0056b3"/>
        <text x="100" y="75" text-anchor="middle" fill="white" font-family="Arial" font-size="12">Markdown</text>
        
        <rect x="200" y="50" width="100" height="40" rx="5" fill="#28a745" stroke="#1e7e34"/>
        <text x="250" y="75" text-anchor="middle" fill="white" font-family="Arial" font-size="12">Processeur</text>
        
        <rect x="350" y="50" width="100" height="40" rx="5" fill="#dc3545" stroke="#c82333"/>
        <text x="400" y="75" text-anchor="middle" fill="white" font-family="Arial" font-size="12">DOCX</text>
        
        <rect x="200" y="150" width="100" height="40" rx="5" fill="#ffc107" stroke="#e0a800"/>
        <text x="250" y="175" text-anchor="middle" fill="#212529" font-family="Arial" font-size="11">Document</text>
        
        <!-- Arrows -->
        <line x1="150" y1="70" x2="200" y2="70" stroke="#333" stroke-width="2" marker-end="url(#arrowhead)"/>
        <line x1="300" y1="70" x2="350" y2="70" stroke="#333" stroke-width="2" marker-end="url(#arrowhead)"/>
        <line x1="250" y1="90" x2="250" y2="150" stroke="#333" stroke-width="2" marker-end="url(#arrowhead)"/>
        
        <!-- Title -->
        <text x="250" y="25" text-anchor="middle" font-family="Arial" font-size="14" font-weight="bold" fill="#333">
          Diagramme de Flux Mermaid
        </text>
      </svg>
    `;
  }

  /**
   * Generate a sequence diagram SVG
   */
  private generateSequenceSVG(code: string): string {
    return `
      <svg width="500" height="300" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <marker id="arrowhead" markerWidth="10" markerHeight="7" 
                  refX="10" refY="3.5" orient="auto">
            <polygon points="0 0, 10 3.5, 0 7" fill="#333" />
          </marker>
        </defs>
        
        <!-- Background -->
        <rect width="500" height="300" fill="#f8f9fa" stroke="#dee2e6" stroke-width="1"/>
        
        <!-- Actors -->
        <rect x="50" y="40" width="80" height="30" fill="#007bff" stroke="#0056b3"/>
        <text x="90" y="60" text-anchor="middle" fill="white" font-family="Arial" font-size="11">Utilisateur</text>
        
        <rect x="200" y="40" width="80" height="30" fill="#28a745" stroke="#1e7e34"/>
        <text x="240" y="60" text-anchor="middle" fill="white" font-family="Arial" font-size="11">Application</text>
        
        <rect x="350" y="40" width="80" height="30" fill="#dc3545" stroke="#c82333"/>
        <text x="390" y="60" text-anchor="middle" fill="white" font-family="Arial" font-size="11">Serveur</text>
        
        <!-- Lifelines -->
        <line x1="90" y1="70" x2="90" y2="250" stroke="#666" stroke-width="2" stroke-dasharray="3,3"/>
        <line x1="240" y1="70" x2="240" y2="250" stroke="#666" stroke-width="2" stroke-dasharray="3,3"/>
        <line x1="390" y1="70" x2="390" y2="250" stroke="#666" stroke-width="2" stroke-dasharray="3,3"/>
        
        <!-- Messages -->
        <line x1="90" y1="100" x2="240" y2="100" stroke="#333" stroke-width="2" marker-end="url(#arrowhead)"/>
        <text x="165" y="95" text-anchor="middle" font-family="Arial" font-size="10" fill="#333">Requ??te</text>
        
        <line x1="240" y1="130" x2="390" y2="130" stroke="#333" stroke-width="2" marker-end="url(#arrowhead)"/>
        <text x="315" y="125" text-anchor="middle" font-family="Arial" font-size="10" fill="#333">API Call</text>
        
        <line x1="390" y1="160" x2="240" y2="160" stroke="#333" stroke-width="2" marker-end="url(#arrowhead)"/>
        <text x="315" y="155" text-anchor="middle" font-family="Arial" font-size="10" fill="#333">R??ponse</text>
        
        <line x1="240" y1="190" x2="90" y2="190" stroke="#333" stroke-width="2" marker-end="url(#arrowhead)"/>
        <text x="165" y="185" text-anchor="middle" font-family="Arial" font-size="10" fill="#333">R??sultat</text>
        
        <!-- Title -->
        <text x="250" y="25" text-anchor="middle" font-family="Arial" font-size="14" font-weight="bold" fill="#333">
          Diagramme de S??quence Mermaid
        </text>
      </svg>
    `;
  }

  /**
   * Generate a default diagram SVG
   */
  private generateDefaultDiagramSVG(diagramType: string): string {
    return `
      <svg width="400" height="200" xmlns="http://www.w3.org/2000/svg">
        <rect width="400" height="200" fill="#f8f9fa" stroke="#dee2e6" stroke-width="2"/>
        <circle cx="200" cy="100" r="60" fill="#007bff" stroke="#0056b3" stroke-width="3"/>
        <text x="200" y="100" text-anchor="middle" dominant-baseline="middle" 
              font-family="Arial, sans-serif" font-size="14" font-weight="bold" fill="white">
          ${diagramType}
        </text>
        <text x="200" y="120" text-anchor="middle" dominant-baseline="middle" 
              font-family="Arial, sans-serif" font-size="10" fill="white">
          Diagram
        </text>
        <text x="200" y="170" text-anchor="middle" dominant-baseline="middle" 
              font-family="Arial, sans-serif" font-size="12" fill="#666">
          G??n??r?? par Mermaid Processor
        </text>
      </svg>
    `;
  }

  /**
   * Detect the type of Mermaid diagram
   */
  private detectDiagramType(code: string): string {
    if (code.includes('graph') || code.includes('flowchart')) return 'Flowchart';
    if (code.includes('sequenceDiagram')) return 'Sequence';
    if (code.includes('classDiagram')) return 'Class';
    if (code.includes('stateDiagram')) return 'State';
    if (code.includes('erDiagram')) return 'ER';
    if (code.includes('gantt')) return 'Gantt';
    if (code.includes('pie')) return 'Pie Chart';
    return 'Mermaid';
  }

  /**
   * Initialize Puppeteer browser
   */
  async initialize(): Promise<void> {
    if (this.isInitialized) {
      return;
    }

    try {
      logger.info('Initializing Mermaid processor...');
      
      this.browser = await puppeteer.launch({
        headless: 'new',
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage',
          '--disable-accelerated-2d-canvas',
          '--no-first-run',
          '--no-zygote',
          '--disable-gpu',
        ],
      });

      this.page = await this.browser.newPage();
      await this.page.setViewport({ width: 1200, height: 800 });

      // Load Mermaid library
      await this.page.addScriptTag({
        url: 'https://cdn.jsdelivr.net/npm/mermaid@10.6.1/dist/mermaid.min.js',
      });

      // Initialize Mermaid with configuration
      await this.page.evaluate((theme) => {
        (window as any).mermaid.initialize({
          startOnLoad: false,
          theme: theme,
          securityLevel: 'loose',
          fontFamily: 'Arial, sans-serif',
          fontSize: 14,
          flowchart: {
            useMaxWidth: true,
            htmlLabels: true,
          },
          sequence: {
            diagramMarginX: 50,
            diagramMarginY: 10,
            actorMargin: 50,
            width: 150,
            height: 65,
            boxMargin: 10,
            boxTextMargin: 5,
            noteMargin: 10,
            messageMargin: 35,
          },
          gantt: {
            gridLineStartPadding: 350,
            fontSize: 11,
            fontFamily: 'Arial, sans-serif',
          },
        });
      }, this.theme);

      this.isInitialized = true;
      logger.info('Mermaid processor initialized successfully');
    } catch (error) {
      logger.error('Failed to initialize Mermaid processor', { error });
      throw error;
    }
  }

  /**
   * Process Mermaid diagram and return SVG/PNG
   */
  async processDiagram(
    mermaidCode: string,
    format: 'svg' | 'png' = 'svg',
    options: {
      width?: number;
      height?: number;
      backgroundColor?: string;
      scale?: number;
    } = {}
  ): Promise<ProcessedMermaidDiagram> {
    if (!this.isInitialized) {
      await this.initialize();
    }

    const {
      width = 800,
      height = 600,
      backgroundColor = 'white',
      scale = 1,
    } = options;

    try {
      logger.debug('Processing Mermaid diagram', { format, width, height });
      
      const { result: diagram, elapsed } = await PerformanceUtils.measureAsync(async () => {
        return await this.renderDiagram(mermaidCode, format, {
          width,
          height,
          backgroundColor,
          scale,
        });
      }, 'mermaid-render');

      logger.debug(`Mermaid diagram processed in ${elapsed}ms`);
      return diagram;
    } catch (error) {
      logger.error('Failed to process Mermaid diagram', { error, code: mermaidCode });
      throw error;
    }
  }

  /**
   * Extract all Mermaid diagrams from Markdown content
   */
  extractMermaidDiagrams(markdown: string): Array<{
    code: string;
    startLine: number;
    endLine: number;
    id: string;
  }> {
    const diagrams: Array<{
      code: string;
      startLine: number;
      endLine: number;
      id: string;
    }> = [];

    const lines = markdown.split('\n');
    let inMermaidBlock = false;
    let currentDiagram: string[] = [];
    let startLine = 0;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      if (line === '```mermaid' || line.startsWith('```mermaid ')) {
        inMermaidBlock = true;
        startLine = i + 1;
        currentDiagram = [];
      } else if (line === '```' && inMermaidBlock) {
        inMermaidBlock = false;
        if (currentDiagram.length > 0) {
          diagrams.push({
            code: currentDiagram.join('\n').trim(),
            startLine,
            endLine: i - 1,
            id: uuidv4(),
          });
        }
        currentDiagram = [];
      } else if (inMermaidBlock) {
        currentDiagram.push(lines[i]);
      }
    }

    logger.info(`Extracted ${diagrams.length} Mermaid diagrams`);
    return diagrams;
  }

  /**
   * Replace Mermaid code blocks with image references
   */
  replaceMermaidBlocks(
    markdown: string,
    processedDiagrams: ProcessedMermaidDiagram[],
    imageBasePath: string = './images'
  ): string {
    const diagrams = this.extractMermaidDiagrams(markdown);
    let updatedMarkdown = markdown;

    // Process in reverse order to maintain line positions
    for (let i = diagrams.length - 1; i >= 0; i--) {
      const diagram = diagrams[i];
      const processedDiagram = processedDiagrams.find(pd => pd.code === diagram.code);

      if (processedDiagram) {
        const imageFileName = `mermaid_${processedDiagram.id}.${processedDiagram.format}`;
        const imagePath = `${imageBasePath}/${imageFileName}`;
        const imageMarkdown = `![Mermaid Diagram](${imagePath})`;

        // Replace the entire mermaid block
        const lines = updatedMarkdown.split('\n');
        const beforeLines = lines.slice(0, diagram.startLine - 1);
        const afterLines = lines.slice(diagram.endLine + 2);
        
        updatedMarkdown = [
          ...beforeLines,
          imageMarkdown,
          ...afterLines,
        ].join('\n');
      }
    }

    return updatedMarkdown;
  }

  /**
   * Validate Mermaid diagram syntax
   */
  async validateDiagram(mermaidCode: string): Promise<{
    isValid: boolean;
    error?: string;
  }> {
    if (!this.isInitialized) {
      await this.initialize();
    }

    try {
      await this.page!.evaluate((code) => {
        return new Promise((resolve, reject) => {
          try {
            (window as any).mermaid.parse(code);
            resolve(true);
          } catch (error) {
            reject(error);
          }
        });
      }, mermaidCode);

      return { isValid: true };
    } catch (error) {
      return {
        isValid: false,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }

  /**
   * Get supported diagram types
   */
  getSupportedTypes(): string[] {
    return [
      'flowchart',
      'graph',
      'sequenceDiagram',
      'classDiagram',
      'stateDiagram',
      'erDiagram',
      'journey',
      'gantt',
      'pie',
      'gitgraph',
      'mindmap',
      'timeline',
      'quadrantChart',
    ];
  }

  /**
   * Cleanup resources
   */
  async cleanup(): Promise<void> {
    try {
      if (this.page) {
        await this.page.close();
        this.page = null;
      }

      if (this.browser) {
        await this.browser.close();
        this.browser = null;
      }

      this.isInitialized = false;
      logger.info('Mermaid processor cleaned up');
    } catch (error) {
      logger.error('Error during Mermaid processor cleanup', { error });
    }
  }

  /**
   * Private method to render diagram
   */
  private async renderDiagram(
    mermaidCode: string,
    format: 'svg' | 'png',
    options: {
      width: number;
      height: number;
      backgroundColor: string;
      scale: number;
    }
  ): Promise<ProcessedMermaidDiagram> {
    const id = uuidv4();

    // Set page content with Mermaid diagram
    await this.page!.setContent(`
      <!DOCTYPE html>
      <html>
        <head>
          <style>
            body {
              margin: 0;
              padding: 20px;
              background-color: ${options.backgroundColor};
              font-family: Arial, sans-serif;
            }
            .mermaid {
              display: flex;
              justify-content: center;
              align-items: center;
              min-height: 100vh;
            }
          </style>
        </head>
        <body>
          <div class="mermaid">${mermaidCode}</div>
          <script src="https://cdn.jsdelivr.net/npm/mermaid@10.6.1/dist/mermaid.min.js"></script>
          <script>
            mermaid.initialize({
              startOnLoad: true,
              theme: '${this.theme}',
              securityLevel: 'loose',
            });
          </script>
        </body>
      </html>
    `);

    // Wait for Mermaid to render
    await this.page!.waitForSelector('.mermaid svg', { timeout: 10000 });

    // Get the SVG element
    const svgElement = await this.page!.$('.mermaid svg');
    if (!svgElement) {
      throw new Error('Failed to render Mermaid diagram');
    }

    // Get dimensions
    const boundingBox = await svgElement.boundingBox();
    const dimensions = {
      width: boundingBox?.width || options.width,
      height: boundingBox?.height || options.height,
    };

    let imageBuffer: Buffer;

    if (format === 'svg') {
      // Get SVG content
      const svgContent = await this.page!.evaluate(() => {
        const svg = document.querySelector('.mermaid svg');
        return svg ? svg.outerHTML : '';
      });

      imageBuffer = Buffer.from(svgContent, 'utf-8');
    } else {
      // Generate PNG
      imageBuffer = await svgElement.screenshot({
        type: 'png',
        omitBackground: options.backgroundColor === 'transparent',
      });
    }

    return {
      code: mermaidCode,
      imageBuffer,
      format,
      dimensions,
      id,
    };
  }
}

/**
 * Mermaid theme configurations
 */
export const MermaidThemes: Record<MermaidTheme, Record<string, any>> = {
  default: {
    primaryColor: '#fff2cc',
    primaryTextColor: '#333',
    primaryBorderColor: '#d6b656',
    lineColor: '#333333',
    secondaryColor: '#d4edda',
    tertiaryColor: '#f8f9fa',
  },
  forest: {
    primaryColor: '#cde498',
    primaryTextColor: '#333',
    primaryBorderColor: '#9cb86f',
    lineColor: '#333333',
    secondaryColor: '#cdffb2',
    tertiaryColor: '#ffffff',
  },
  dark: {
    primaryColor: '#1f2328',
    primaryTextColor: '#f0f6fc',
    primaryBorderColor: '#30363d',
    lineColor: '#f0f6fc',
    secondaryColor: '#262c36',
    tertiaryColor: '#161b22',
  },
  neutral: {
    primaryColor: '#eaeaea',
    primaryTextColor: '#333',
    primaryBorderColor: '#cccccc',
    lineColor: '#333333',
    secondaryColor: '#f5f5f5',
    tertiaryColor: '#ffffff',
  },
  base: {
    primaryColor: '#ececff',
    primaryTextColor: '#333',
    primaryBorderColor: '#9370db',
    lineColor: '#333333',
    secondaryColor: '#e0e0e0',
    tertiaryColor: '#ffffff',
  },
};
