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
    const textParts: string[] = [];

    const extractTextFromElement = (element: any) => {
      if (!element) return;

      // Handle text content directly - this is the actual readable text
      if (typeof element === 'string') {
        const cleanText = element.trim();
        if (cleanText && !cleanText.includes('http://schemas.microsoft.com')) {
          textParts.push(cleanText);
        }
        return;
      }

      // Handle arrays
      if (Array.isArray(element)) {
        element.forEach(item => extractTextFromElement(item));
        return;
      }

      // Handle objects
      if (typeof element === 'object') {
        // Text content - the actual readable text in Word documents
        if (element['w:t']) {
          let text = '';
          
          // Handle direct text content
          if (typeof element['w:t'] === 'string') {
            text = element['w:t'];
          } else if (element['w:t'] && typeof element['w:t'] === 'object' && element['w:t']['_']) {
            // Handle complex text objects with underscore property
            text = element['w:t']['_'];
          } else if (element['w:t'] && typeof element['w:t'] === 'object' && element['w:t']['$']) {
            // Handle text with attributes
            text = element['w:t']['$']['xml:space'] === 'preserve' ? element['w:t']['_'] || '' : '';
          }
          
          // Only add meaningful text (not XML garbage)
          if (text && typeof text === 'string' && text.trim() && !text.includes('http://schemas')) {
            textParts.push(text);
          }
          return;
        }

        // Paragraph break
        if (element['w:p']) {
          if (textParts.length > 0 && !textParts[textParts.length - 1].endsWith('\n')) {
            textParts.push('\n\n');
          }
          extractTextFromElement(element['w:p']);
          return;
        }

        // Run (text formatting container)
        if (element['w:r']) {
          extractTextFromElement(element['w:r']);
          return;
        }

        // Tab character
        if (element['w:tab']) {
          textParts.push('\t');
          return;
        }

        // Line break
        if (element['w:br']) {
          textParts.push('\n');
          return;
        }

        // Document structure elements
        if (element['w:document']) {
          extractTextFromElement(element['w:document']);
          return;
        }

        if (element['w:body']) {
          extractTextFromElement(element['w:body']);
          return;
        }

        // Recursively process other Word elements, but avoid formatting/style elements
        for (const [key, value] of Object.entries(element)) {
          if (key.startsWith('w:') && 
              !key.includes('Pr') && 
              !key.includes('Style') && 
              !key.includes('Properties') &&
              !key.includes('Fonts') &&
              !key.includes('Settings') &&
              key !== 'w:t') {
            extractTextFromElement(value);
          }
        }
      }
    };

    extractTextFromElement(parsedDoc);
    
    // Clean up and format the extracted text
    let finalText = textParts
      .join(' ')  // Join with spaces instead of empty string
      .replace(/\n{3,}/g, '\n\n')     // Max 2 consecutive newlines
      .replace(/[ \t]+/g, ' ')        // Multiple spaces/tabs to single space
      .replace(/\n /g, '\n')          // Remove spaces after newlines
      .replace(/([a-z])([A-Z])/g, '$1 $2')  // Add space between camelCase words
      .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')  // Add space between consecutive caps
      .replace(/([0-9])([A-Z])/g, '$1 $2')  // Add space between numbers and caps
      .replace(/([a-z])([0-9])/g, '$1 $2')  // Add space between lowercase and numbers
      .replace(/[0-9A-F]{8,}/g, '')   // Remove hex codes (8+ hex chars)
      .replace(/\s+/g, ' ')           // Normalize whitespace
      .trim();

    // If we got no meaningful text, try a fallback approach
    if (!finalText || finalText.length < 10) {
      // Fallback: extract any readable text from the document
      const fallbackText = this.extractTextFallback(parsedDoc);
      if (fallbackText && fallbackText.length > finalText.length) {
        finalText = fallbackText;
      }
    }

    return finalText;
  }

  private extractTextFallback(obj: any): string {
    const textParts: string[] = [];
    
    const findText = (element: any) => {
      if (typeof element === 'string' && element.trim() && !element.includes('http://schemas')) {
        textParts.push(element.trim());
      } else if (Array.isArray(element)) {
        element.forEach(findText);
      } else if (element && typeof element === 'object') {
        Object.values(element).forEach(findText);
      }
    };
    
    findText(obj);
    
    return textParts
      .filter(text => text.length > 2 && !/^[0-9]+$/.test(text)) // Filter out single chars and pure numbers
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();
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

    // Better paragraph detection and formatting
    const sentences = content.split(/(?<=[.!?])\s+/);
    let currentParagraph = '';
    
    for (const sentence of sentences) {
      const cleanSentence = sentence.trim();
      if (!cleanSentence) continue;
      
      // If this looks like a header (starts with capital and is short)
      if (cleanSentence.length < 50 && /^[A-Z][a-z]*\s*[A-Z]/.test(cleanSentence)) {
        if (currentParagraph) {
          markdown += currentParagraph.trim() + '\n\n';
          currentParagraph = '';
        }
        markdown += `## ${cleanSentence}\n\n`;
      } else {
        currentParagraph += cleanSentence + ' ';
        
        // End paragraph if we have enough content
        if (currentParagraph.length > 200) {
          markdown += currentParagraph.trim() + '\n\n';
          currentParagraph = '';
        }
      }
    }
    
    // Add any remaining content
    if (currentParagraph.trim()) {
      markdown += currentParagraph.trim() + '\n\n';
    }

    return markdown.trim();
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