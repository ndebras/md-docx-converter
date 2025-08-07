"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MermaidPNGProcessor = void 0;
const puppeteer_1 = __importDefault(require("puppeteer"));
const fs = __importStar(require("fs-extra"));
const path = __importStar(require("path"));
const uuid_1 = require("uuid");
const sharp_1 = __importDefault(require("sharp"));
/**
 * Nouveau processeur Mermaid avec g??n??ration PNG compl??te
 */
class MermaidPNGProcessor {
    constructor() {
        this.browser = null;
        this.tempDir = path.join(process.cwd(), 'temp-mermaid');
        this.ensureTempDir();
    }
    /**
     * Traite le contenu et remplace les diagrammes Mermaid par des images PNG
     */
    async processContent(content) {
        const mermaidBlocks = content.match(/```mermaid[\s\S]*?```/g) || [];
        if (mermaidBlocks.length === 0) {
            return { content, diagramCount: 0, images: [] };
        }
        const images = [];
        let processedContent = content;
        let diagramCount = 0;
        // Initialiser Puppeteer
        await this.initBrowser();
        for (const block of mermaidBlocks) {
            try {
                const diagramCode = block.replace(/```mermaid\s*\n/, '').replace(/\n```$/, '').trim();
                if (diagramCode) {
                    const imageId = (0, uuid_1.v4)();
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
            }
            catch (error) {
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
    async generateMermaidPNG(diagramCode, id) {
        if (!this.browser) {
            throw new Error('Browser non initialis??');
        }
        try {
            const page = await this.browser.newPage();
            await page.setViewport({ width: 1200, height: 800 });
            // HTML pour rendre le diagramme Mermaid
            const html = `
<!DOCTYPE html>
<html>
<head>
    <script src="https://cdn.jsdelivr.net/npm/mermaid@10.6.1/dist/mermaid.min.js"></script>
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
            securityLevel: 'loose',
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
            const optimizedBuffer = await (0, sharp_1.default)(screenshot)
                .png({ quality: 90, compressionLevel: 6 })
                .resize({
                width: Math.round(Math.min(800, boundingBox.width)),
                height: Math.round(Math.min(600, boundingBox.height)),
                fit: 'inside',
                withoutEnlargement: true
            })
                .toBuffer();
            return optimizedBuffer;
        }
        catch (error) {
            console.error(`Erreur g??n??ration PNG pour ${id}:`, error);
            return null;
        }
    }
    /**
     * Initialise le navigateur Puppeteer
     */
    async initBrowser() {
        if (!this.browser) {
            this.browser = await puppeteer_1.default.launch({
                headless: 'new',
                args: ['--no-sandbox', '--disable-setuid-sandbox']
            });
        }
    }
    /**
     * Nettoie les ressources
     */
    async cleanup() {
        if (this.browser) {
            await this.browser.close();
            this.browser = null;
        }
        // Nettoyer les fichiers temporaires
        try {
            if (await fs.pathExists(this.tempDir)) {
                await fs.emptyDir(this.tempDir);
            }
        }
        catch (error) {
            console.warn('Erreur lors du nettoyage des fichiers temporaires:', error);
        }
    }
    /**
     * Assure que le r??pertoire temporaire existe
     */
    ensureTempDir() {
        if (!fs.existsSync(this.tempDir)) {
            fs.mkdirSync(this.tempDir, { recursive: true });
        }
    }
    /**
     * Nettoie le r??pertoire temporaire
     */
    cleanupTempDir() {
        if (fs.existsSync(this.tempDir)) {
            fs.rmSync(this.tempDir, { recursive: true, force: true });
        }
    }
}
exports.MermaidPNGProcessor = MermaidPNGProcessor;
//# sourceMappingURL=mermaid-png-processor.js.map