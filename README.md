# FSS Parse Word TypeScript

Professional TypeScript Word document parsing and manipulation toolkit for Node.js applications.

## 🚀 Features

- **Multiple Format Support**: DOCX, DOC, RTF, ODT
- **Flexible Output**: Text, Markdown, HTML, JSON
- **Metadata Extraction**: Complete document properties
- **Image Extraction**: Extract embedded images
- **Safety First**: Built-in malware scanning and validation
- **TypeScript**: Full type safety and IntelliSense support
- **CLI & Library**: Use as command-line tool or Node.js library
- **Performance**: Optimized for large documents
- **Enterprise Ready**: Production-grade error handling

## 📦 Installation

### Global Installation (CLI)
```bash
npm install -g fss-parse-word-ts
```

### Project Installation (Library)
```bash
npm install fss-parse-word-ts
```

## 🖥️ CLI Usage

### Extract Text
```bash
# Basic text extraction
fss-parse-word-ts extract document.docx

# Extract to file with specific format
fss-parse-word-ts extract document.docx -o output.md -f markdown

# Extract with images
fss-parse-word-ts extract document.docx --images
```

### Convert Documents
```bash
# Convert to markdown
fss-parse-word-ts convert document.docx output.md -f markdown

# Convert to HTML with formatting
fss-parse-word-ts convert document.docx output.html -f html --preserve-formatting

# Convert to JSON with metadata
fss-parse-word-ts convert document.docx output.json -f json
```

### Document Information
```bash
# Basic info
fss-parse-word-ts info document.docx

# Detailed analysis
fss-parse-word-ts info document.docx --detailed
```

### Validate Documents
```bash
# Security and integrity validation
fss-parse-word-ts validate document.docx
```

## 📚 Library Usage

### Basic Parsing
```typescript
import { WordParser } from 'fss-parse-word-ts';

const parser = new WordParser({
  includeMetadata: true,
  extractImages: false,
  outputFormat: 'text'
});

const result = await parser.parseDocument('document.docx');
if (result.success) {
  console.log('Content:', result.content);
  console.log('Metadata:', result.metadata);
}
```

### Advanced Configuration
```typescript
import { WordParser, WordConfig } from 'fss-parse-word-ts';

const config: WordConfig = {
  extractImages: true,
  preserveFormatting: true,
  includeMetadata: true,
  outputFormat: 'markdown',
  safetyChecks: true
};

const parser = new WordParser(config);
const result = await parser.parseDocument('document.docx');

// Convert to different format
const markdown = await parser.convertToFormat(
  result.content!,
  'markdown',
  result.metadata
);
```

### Safety Management
```typescript
import { SafetyManager } from 'fss-parse-word-ts';

const safety = new SafetyManager();
const validation = await safety.validateFile('document.docx');

if (validation.isSafe) {
  // Create backup before processing
  const backupPath = await safety.createBackup('document.docx');
  console.log('Backup created at:', backupPath);
} else {
  console.log('Safety issues:', validation.issues);
}
```

## 🔧 Configuration Options

```typescript
interface WordConfig {
  extractImages?: boolean;        // Extract embedded images
  preserveFormatting?: boolean;   // Preserve text formatting
  includeMetadata?: boolean;      // Include document metadata
  outputFormat?: 'text' | 'markdown' | 'html' | 'json';
  safetyChecks?: boolean;         // Enable safety validation
}
```

## 📊 Output Formats

### Text
Clean plain text with paragraph breaks.

### Markdown
Structured markdown with headings, formatting, and metadata.

### HTML
Complete HTML document with proper structure.

### JSON
Structured data with content, metadata, and processing information.

## 🛡️ Security Features

- **File Validation**: Extension and size checks
- **Malware Scanning**: Basic suspicious pattern detection
- **Backup System**: Automatic backup creation before processing
- **Hash Verification**: File integrity checking
- **Size Limits**: Configurable maximum file size

## 🔧 Development

### Setup
```bash
git clone https://github.com/FSSCoding/fss-parse-word-ts.git
cd fss-parse-word-ts
npm install
```

### Build
```bash
npm run build
```

### Testing
```bash
npm test
npm run test:watch
```

### Linting
```bash
npm run lint
npm run lint:fix
```

## 🏗️ Architecture

- **WordParser**: Main parsing engine with multi-format support
- **SafetyManager**: Security validation and backup management
- **CLI**: Command-line interface with rich output
- **Types**: Complete TypeScript type definitions

## 📋 Requirements

- Node.js 16.0.0 or higher
- TypeScript 5.0+ (for development)

## 📄 License

MIT License - see LICENSE file for details.

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch
3. Add tests for new functionality
4. Ensure all tests pass
5. Submit a pull request

## 🆘 Support

- 🐛 [Report Issues](https://github.com/FSSCoding/fss-parse-word-ts/issues)
- 📖 [Documentation](https://github.com/FSSCoding/fss-parse-word-ts#readme)
- 💬 [Discussions](https://github.com/FSSCoding/fss-parse-word-ts/discussions)

## 🔗 Related Projects

- **fss-parse-excel-ts** - Excel/spreadsheet parsing
- **fss-parse-pdf-ts** - PDF document processing
- **fss-parse-word** - Python version

---

**Built with ❤️ by FSS Coding - Professional document processing solutions**