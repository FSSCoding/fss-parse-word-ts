import * as fs from 'fs';
import * as crypto from 'crypto';
import * as path from 'path';
import { SafetyResult } from './types';

export class SafetyManager {
  private maxFileSize: number;
  private allowedExtensions: Set<string>;
  private backupDir: string;

  constructor() {
    this.maxFileSize = 100 * 1024 * 1024; // 100MB
    this.allowedExtensions = new Set(['.docx', '.doc', '.rtf', '.odt']);
    this.backupDir = path.join(process.cwd(), '.fss-word-backups');
  }

  async validateFile(filePath: string): Promise<SafetyResult> {
    const issues: string[] = [];

    try {
      // Check file existence
      if (!fs.existsSync(filePath)) {
        issues.push('File does not exist');
        return { isSafe: false, issues, hash: '', fileSize: 0 };
      }

      // Check file extension
      const ext = path.extname(filePath).toLowerCase();
      if (!this.allowedExtensions.has(ext)) {
        issues.push(`Unsupported file extension: ${ext}`);
      }

      // Get file stats
      const stats = fs.statSync(filePath);
      const fileSize = stats.size;

      // Check file size
      if (fileSize > this.maxFileSize) {
        issues.push(`File too large: ${fileSize} bytes (max: ${this.maxFileSize})`);
      }

      // Calculate file hash
      const fileBuffer = fs.readFileSync(filePath);
      const hash = crypto.createHash('sha256').update(fileBuffer).digest('hex');

      // Basic malware checks
      if (await this.basicMalwareCheck(fileBuffer)) {
        issues.push('Potential security risk detected');
      }

      return {
        isSafe: issues.length === 0,
        issues,
        hash,
        fileSize
      };

    } catch (error) {
      issues.push(`Validation error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      return { isSafe: false, issues, hash: '', fileSize: 0 };
    }
  }

  async createBackup(filePath: string): Promise<string> {
    try {
      // Ensure backup directory exists
      if (!fs.existsSync(this.backupDir)) {
        fs.mkdirSync(this.backupDir, { recursive: true });
      }

      const fileName = path.basename(filePath);
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const backupPath = path.join(this.backupDir, `${timestamp}-${fileName}`);

      fs.copyFileSync(filePath, backupPath);
      return backupPath;

    } catch (error) {
      throw new Error(`Backup failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  async cleanupBackups(maxAge: number = 7): Promise<number> {
    try {
      if (!fs.existsSync(this.backupDir)) {
        return 0;
      }

      const files = fs.readdirSync(this.backupDir);
      const cutoffTime = Date.now() - (maxAge * 24 * 60 * 60 * 1000);
      let deletedCount = 0;

      for (const file of files) {
        const filePath = path.join(this.backupDir, file);
        const stats = fs.statSync(filePath);

        if (stats.mtime.getTime() < cutoffTime) {
          fs.unlinkSync(filePath);
          deletedCount++;
        }
      }

      return deletedCount;

    } catch (error) {
      throw new Error(`Cleanup failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  private async basicMalwareCheck(buffer: Buffer): Promise<boolean> {
    // Basic suspicious pattern detection
    const suspiciousPatterns = [
      'eval(',
      'exec(',
      'system(',
      'shell_exec(',
      'base64_decode(',
      'javascript:',
      '<script',
      'vbscript:'
    ];

    const content = buffer.toString('utf8', 0, Math.min(buffer.length, 10000));
    
    return suspiciousPatterns.some(pattern => 
      content.toLowerCase().includes(pattern.toLowerCase())
    );
  }

  setMaxFileSize(size: number): void {
    this.maxFileSize = size;
  }

  addAllowedExtension(ext: string): void {
    this.allowedExtensions.add(ext.toLowerCase());
  }

  setBackupDirectory(dir: string): void {
    this.backupDir = dir;
  }
}