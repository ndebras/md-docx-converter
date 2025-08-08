import puppeteer, { Browser, Page } from 'puppeteer';
import * as fs from 'fs-extra';
import * as path from 'path';
import { v4 as uuidv4 } from 'uuid';
import sharp from 'sharp';
import mermaid from 'mermaid';

/**
 * Nouveau processeur Mermaid avec g??n??ration PNG compl??te
 */
export class MermaidPNGProcessor {
  private browser: Browser | null = null;
  private tempDir: string;

  constructor() {
    this.tempDir = path.join(process.cwd(), 'temp-mermaid');
    this.ensureTempDir();
  }

  /**
   * Traite le contenu et remplace les diagrammes Mermaid par des images PNG
   */
  async processContent(content: string): Promise<{ content: string; diagramCount: number; images: Array<{ id: string; path: string; buffer: Buffer }> }> {
    const mermaidBlocks = content.match(/```mermaid[\s\S]*?```/g) || [];
    
    if (mermaidBlocks.length === 0) {
      return { content, diagramCount: 0, images: [] };
    }

    const images: Array<{ id: string; path: string; buffer: Buffer }> = [];
    let processedContent = content;
    let diagramCount = 0;

    // Initialiser Puppeteer
    await this.initBrowser();

    for (const block of mermaidBlocks) {
      try {
        const diagramCode = block.replace(/```mermaid\s*\n/, '').replace(/\n```$/, '').trim();
        
        if (diagramCode) {
          const imageId = uuidv4();
          const pngBuffer = await this.generateMermaidPNG(diagramCode, imageId);
          
          if (pngBuffer) {
            const imagePath = path.join(this.tempDir, `${imageId}.png`);
            fs.writeFileSync(imagePath, pngBuffer);
            
            images.push({
              id: imageId,
              path: imagePath,
              buffer: pngBuffer
            });

            // Remplacer par une r??f??rence d'image
            const imageRef = `![Mermaid Diagram ${imageId}](mermaid-${imageId}.png)`;
            processedContent = processedContent.replace(block, imageRef);
            diagramCount++;
          }
        }
      } catch (error) {
        console.warn('Erreur lors du traitement du diagramme Mermaid:', error);
        // Garder le bloc original en cas d'??chec
      }
    }

    await this.cleanup();

    return {
      content: processedContent,
      diagramCount,
      images
    };
  }

  /**
   * G??n??re une image PNG ?? partir du code Mermaid
   */
  private async generateMermaidPNG(diagramCode: string, id: string): Promise<Buffer | null> {
    if (!this.browser) {
      throw new Error('Browser non initialis??');
    }

    try {
      const page = await this.browser.newPage();
      await page.setViewport({ width: 1200, height: 800 });
      await page.setOfflineMode(true);

      const mermaidPath = require.resolve('mermaid/dist/mermaid.min.js');

      // HTML pour rendre le diagramme Mermaid avec le script local
      const html = `
<!DOCTYPE html>
<html>
<head>
    <script src="file://${mermaidPath}"></script>
    <style>
        body {
            margin: 0;
            padding: 20px;
            font-family: Arial, sans-serif;
            background: white;
        }
        .mermaid {
            text-align: center;
        }
    </style>
</head>
<body>
    <div class="mermaid">
${diagramCode}
    </div>
    <script>
        mermaid.initialize({
            startOnLoad: true,
            theme: 'default',
            securityLevel: 'strict',
            flowchart: {
                htmlLabels: true,
                curve: 'basis'
            }
        });
    </script>
</body>
</html>`;

      await page.setContent(html);
      
      // Attendre que Mermaid soit rendu
      await page.waitForSelector('.mermaid svg', { timeout: 10000 });
      
      // Obtenir les dimensions du SVG
      const svgElement = await page.$('.mermaid svg');
      if (!svgElement) {
        throw new Error('SVG non trouv??');
      }

      const boundingBox = await svgElement.boundingBox();
      if (!boundingBox) {
        throw new Error('Impossible d\'obtenir les dimensions');
      }

      // Capturer le screenshot de la zone SVG
      const screenshot = await svgElement.screenshot({
        type: 'png',
        omitBackground: false
      });

      await page.close();

      // Optimiser l'image avec Sharp
      const optimizedBuffer = await sharp(screenshot)
        .png({ quality: 90, compressionLevel: 6 })
        .resize({ 
          width: Math.round(Math.min(800, boundingBox.width)), 
          height: Math.round(Math.min(600, boundingBox.height)),
          fit: 'inside',
          withoutEnlargement: true
        })
        .toBuffer();

      return optimizedBuffer;

    } catch (error) {
      console.error(`Erreur g??n??ration PNG pour ${id}:`, error);
      return null;
    }
  }

  /**
   * Initialise le navigateur Puppeteer
   */
  private async initBrowser(): Promise<void> {
    if (!this.browser) {
      this.browser = await puppeteer.launch({
        headless: 'new',
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--allow-file-access-from-files']
      });
    }
  }

  /**
   * Nettoie les ressources
   */
  async cleanup(): Promise<void> {
    if (this.browser) {
      await this.browser.close();
      this.browser = null;
    }
    
    // Nettoyer les fichiers temporaires
    try {
      if (await fs.pathExists(this.tempDir)) {
        await fs.emptyDir(this.tempDir);
      }
    } catch (error) {
      console.warn('Erreur lors du nettoyage des fichiers temporaires:', error);
    }
  }

  /**
   * Assure que le r??pertoire temporaire existe
   */
  private ensureTempDir(): void {
    if (!fs.existsSync(this.tempDir)) {
      fs.mkdirSync(this.tempDir, { recursive: true });
    }
  }

  /**
   * Nettoie le r??pertoire temporaire
   */
  cleanupTempDir(): void {
    if (fs.existsSync(this.tempDir)) {
      fs.rmSync(this.tempDir, { recursive: true, force: true });
    }
  }
}
