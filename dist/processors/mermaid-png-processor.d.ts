/**
 * Nouveau processeur Mermaid avec g??n??ration PNG compl??te
 */
export declare class MermaidPNGProcessor {
    private browser;
    private tempDir;
    constructor();
    /**
     * Traite le contenu et remplace les diagrammes Mermaid par des images PNG
     */
    processContent(content: string): Promise<{
        content: string;
        diagramCount: number;
        images: Array<{
            id: string;
            path: string;
            buffer: Buffer;
        }>;
    }>;
    /**
     * G??n??re une image PNG ?? partir du code Mermaid
     */
    private generateMermaidPNG;
    /**
     * Initialise le navigateur Puppeteer
     */
    private initBrowser;
    /**
     * Nettoie les ressources
     */
    cleanup(): Promise<void>;
    /**
     * Assure que le r??pertoire temporaire existe
     */
    private ensureTempDir;
    /**
     * Nettoie le r??pertoire temporaire
     */
    cleanupTempDir(): void;
}
//# sourceMappingURL=mermaid-png-processor.d.ts.map