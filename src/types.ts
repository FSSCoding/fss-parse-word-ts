export interface WordConfig {
  extractImages?: boolean;
  preserveFormatting?: boolean;
  includeMetadata?: boolean;
  outputFormat?: 'text' | 'markdown' | 'html' | 'json';
  safetyChecks?: boolean;
}

export interface DocumentMetadata {
  title?: string;
  author?: string;
  subject?: string;
  creator?: string;
  created?: Date;
  modified?: Date;
  pages?: number;
  words?: number;
  characters?: number;
}

export interface ProcessingResult {
  success: boolean;
  content?: string;
  metadata?: DocumentMetadata;
  images?: Buffer[];
  warnings?: string[];
  errors?: string[];
  processingTime?: number;
}

export interface SafetyResult {
  isSafe: boolean;
  issues: string[];
  hash: string;
  fileSize: number;
}

export interface ConversionOptions {
  inputPath: string;
  outputPath?: string;
  format: 'text' | 'markdown' | 'html' | 'json';
  config?: WordConfig;
}

export enum LogLevel {
  ERROR = 0,
  WARN = 1,
  INFO = 2,
  DEBUG = 3
}

export interface LogEntry {
  level: LogLevel;
  message: string;
  timestamp: Date;
  context?: any;
}