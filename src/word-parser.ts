import * as yauzl from 'yauzl';
import * as xml2js from 'xml2js';
import { WordConfig, DocumentMetadata, ProcessingResult } from './types';
import { SafetyManager } from './safety-manager';

export class WordParser {
  private config: WordConfig;
  private safetyManager: SafetyManager;
  private parser: xml2js.Parser;

  constructor(config: WordConfig = {}) {
    this.config = {
      extractImages: false,
      preserveFormatting: true,
      includeMetadata: true,
      outputFormat: 'text',
      safetyChecks: true,
      ...config
    };
    
    this.safetyManager = new SafetyManager();
    this.parser = new xml2js.Parser({
      explicitArray: false,
      ignoreAttrs: false,
      mergeAttrs: true
    });
  }

  async parseDocument(filePath: string): Promise<ProcessingResult> {
    const startTime = Date.now();
    const result: ProcessingResult = {
      success: false,
      warnings: [],
      errors: []
    };

    try {
      // Safety validation
      if (this.config.safetyChecks) {
        const safetyResult = await this.safetyManager.validateFile(filePath);
        if (!safetyResult.isSafe) {
          result.errors = safetyResult.issues;
          return result;
        }
      }

      // Extract DOCX content
      const docxContent = await this.extractDocxContent(filePath);
      
      // Parse document structure
      const documentXml = docxContent['word/document.xml'];
      if (!documentXml) {
        result.errors?.push('Invalid DOCX: Missing document.xml');
        return result;
      }

      const parsedDoc = await this.parser.parseStringPromise(documentXml);
      
      // Extract text content
      result.content = this.extractText(parsedDoc);

      // Extract metadata if requested
      if (this.config.includeMetadata) {
        result.metadata = await this.extractMetadata(docxContent);
      }

      // Extract images if requested
      if (this.config.extractImages) {
        result.images = await this.extractImages(docxContent);
      }

      result.success = true;
      result.processingTime = Date.now() - startTime;

    } catch (error) {
      result.errors?.push(`Parsing failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }

    return result;
  }

  private async extractDocxContent(filePath: string): Promise<Record<string, string>> {
    return new Promise((resolve, reject) => {
      const content: Record<string, string> = {};

      yauzl.open(filePath, { lazyEntries: true }, (err, zipfile) => {
        if (err || !zipfile) {
          reject(err || new Error('Failed to open DOCX file'));
          return;
        }

        zipfile.readEntry();

        zipfile.on('entry', (entry) => {
          if (/\/$/.test(entry.fileName)) {
            // Directory entry
            zipfile.readEntry();
          } else {
            // File entry
            zipfile.openReadStream(entry, (err, readStream) => {
              if (err) {
                reject(err);
                return;
              }

              if (!readStream) {
                zipfile.readEntry();
                return;
              }

              const chunks: Buffer[] = [];
              readStream.on('data', (chunk) => chunks.push(chunk));
              readStream.on('end', () => {
                content[entry.fileName] = Buffer.concat(chunks).toString('utf8');
                zipfile.readEntry();
              });
            });
          }
        });

        zipfile.on('end', () => {
          resolve(content);
        });

        zipfile.on('error', (err) => {
          reject(err);
        });
      });
    });
  }

  private extractText(parsedDoc: any): string {
    const texts: string[] = [];

    const extractFromNode = (node: any) => {
      if (typeof node === 'string') {
        texts.push(node);
        return;
      }

      if (Array.isArray(node)) {
        node.forEach(extractFromNode);
        return;
      }

      if (typeof node === 'object' && node !== null) {
        // Handle text nodes
        if (node['w:t']) {
          const textContent = Array.isArray(node['w:t']) ? node['w:t'].join('') : node['w:t'];
          texts.push(textContent);
        }

        // Handle paragraphs and runs
        if (node['w:p']) {
          const paragraphs = Array.isArray(node['w:p']) ? node['w:p'] : [node['w:p']];
          paragraphs.forEach(extractFromNode);
          texts.push('\n');
        }

        if (node['w:r']) {
          const runs = Array.isArray(node['w:r']) ? node['w:r'] : [node['w:r']];
          runs.forEach(extractFromNode);
        }

        // Recursively process other nodes
        Object.values(node).forEach(extractFromNode);
      }
    };

    extractFromNode(parsedDoc);
    return texts.join('').replace(/\n+/g, '\n').trim();
  }

  private async extractMetadata(docxContent: Record<string, string>): Promise<DocumentMetadata> {
    const metadata: DocumentMetadata = {};

    try {
      // Core properties
      const coreProps = docxContent['docProps/core.xml'];
      if (coreProps) {
        const parsed = await this.parser.parseStringPromise(coreProps);
        const props = parsed['cp:coreProperties'] || {};
        
        metadata.title = props['dc:title'];
        metadata.author = props['dc:creator'];
        metadata.subject = props['dc:subject'];
        if (props['dcterms:created']) metadata.created = new Date(props['dcterms:created']);
        if (props['dcterms:modified']) metadata.modified = new Date(props['dcterms:modified']);
      }

      // App properties  
      const appProps = docxContent['docProps/app.xml'];
      if (appProps) {
        const parsed = await this.parser.parseStringPromise(appProps);
        const props = parsed.Properties || {};
        
        if (props.Pages) metadata.pages = parseInt(props.Pages);
        if (props.Words) metadata.words = parseInt(props.Words);
        if (props.Characters) metadata.characters = parseInt(props.Characters);
      }

    } catch (error) {
      // Metadata extraction is optional, don't fail the whole process
    }

    return metadata;
  }

  private async extractImages(docxContent: Record<string, string>): Promise<Buffer[]> {
    const images: Buffer[] = [];

    try {
      // Find image files in the DOCX
      const imageFiles = Object.keys(docxContent).filter(fileName =>
        fileName.startsWith('word/media/') && /\.(png|jpg|jpeg|gif|bmp)$/i.test(fileName)
      );

      for (const imageFile of imageFiles) {
        const imageData = docxContent[imageFile];
        if (imageData) {
          // Convert base64 to buffer if needed
          const buffer = Buffer.isBuffer(imageData) ? imageData : Buffer.from(imageData, 'base64');
          images.push(buffer);
        }
      }

    } catch (error) {
      // Image extraction is optional
    }

    return images;
  }

  async convertToFormat(content: string, format: string, metadata?: DocumentMetadata): Promise<string> {
    switch (format.toLowerCase()) {
      case 'markdown':
        return this.convertToMarkdown(content, metadata);
      case 'html':
        return this.convertToHtml(content, metadata);
      case 'json':
        return JSON.stringify({
          content,
          metadata,
          format: 'json',
          timestamp: new Date().toISOString()
        }, null, 2);
      case 'text':
      default:
        return content;
    }
  }

  private convertToMarkdown(content: string, metadata?: DocumentMetadata): string {
    let markdown = '';

    if (metadata?.title) {
      markdown += `# ${metadata.title}\n\n`;
    }

    if (metadata?.author) {
      markdown += `**Author:** ${metadata.author}\n\n`;
    }

    // Simple paragraph conversion
    const paragraphs = content.split('\n').filter(p => p.trim());
    markdown += paragraphs.join('\n\n');

    return markdown;
  }

  private convertToHtml(content: string, metadata?: DocumentMetadata): string {
    let html = '<!DOCTYPE html>\n<html>\n<head>\n';
    html += '<meta charset="utf-8">\n';
    
    if (metadata?.title) {
      html += `<title>${this.escapeHtml(metadata.title)}</title>\n`;
    }
    
    html += '</head>\n<body>\n';

    if (metadata?.title) {
      html += `<h1>${this.escapeHtml(metadata.title)}</h1>\n`;
    }

    // Convert paragraphs to HTML
    const paragraphs = content.split('\n').filter(p => p.trim());
    paragraphs.forEach(paragraph => {
      html += `<p>${this.escapeHtml(paragraph)}</p>\n`;
    });

    html += '</body>\n</html>';
    return html;
  }

  private escapeHtml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  updateConfig(newConfig: Partial<WordConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }

  getConfig(): WordConfig {
    return { ...this.config };
  }
}