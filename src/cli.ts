#!/usr/bin/env node

import { Command } from 'commander';
import chalk from 'chalk';
import ora from 'ora';
import * as fs from 'fs';
import * as path from 'path';
import { WordParser } from './word-parser';
import { WordConfig } from './types';

const program = new Command();

program
  .name('fss-parse-word-ts')
  .description('Professional TypeScript Word document parsing and manipulation toolkit')
  .version('1.0.0');

program
  .command('extract')
  .description('Extract text from Word document')
  .argument('<input>', 'Input Word document path')
  .option('-o, --output <path>', 'Output file path')
  .option('-f, --format <format>', 'Output format (text|markdown|html|json)', 'text')
  .option('--no-metadata', 'Skip metadata extraction')
  .option('--images', 'Extract images')
  .option('--no-safety', 'Skip safety checks')
  .action(async (input: string, options: any) => {
    const spinner = ora('Extracting content from Word document...').start();

    try {
      const config: WordConfig = {
        includeMetadata: options.metadata !== false,
        extractImages: options.images,
        safetyChecks: options.safety !== false,
        outputFormat: options.format
      };

      const parser = new WordParser(config);
      const result = await parser.parseDocument(input);

      if (!result.success) {
        spinner.fail('Extraction failed');
        if (result.errors?.length) {
          console.error(chalk.red('Errors:'));
          result.errors.forEach(error => console.error(chalk.red(`  • ${error}`)));
        }
        process.exit(1);
      }

      if (result.content) {
        const output = await parser.convertToFormat(result.content, options.format, result.metadata);
        
        if (options.output) {
          fs.writeFileSync(options.output, output, 'utf8');
          spinner.succeed(`Content extracted to ${chalk.green(options.output)}`);
        } else {
          spinner.stop();
          console.log(output);
        }
      }

      // Show metadata if available
      if (result.metadata && options.format !== 'json') {
        console.log(chalk.blue('\nDocument Information:'));
        if (result.metadata.title) console.log(`  Title: ${result.metadata.title}`);
        if (result.metadata.author) console.log(`  Author: ${result.metadata.author}`);
        if (result.metadata.pages) console.log(`  Pages: ${result.metadata.pages}`);
        if (result.metadata.words) console.log(`  Words: ${result.metadata.words}`);
      }

      // Show processing time
      if (result.processingTime) {
        console.log(chalk.gray(`\nProcessed in ${result.processingTime}ms`));
      }

    } catch (error) {
      spinner.fail('Processing failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

program
  .command('convert')
  .description('Convert Word document to different formats')
  .argument('<input>', 'Input Word document path')
  .argument('<output>', 'Output file path')
  .option('-f, --format <format>', 'Output format (text|markdown|html|json)', 'markdown')
  .option('--preserve-formatting', 'Preserve text formatting')
  .option('--images', 'Extract and include images')
  .action(async (input: string, output: string, options: any) => {
    const spinner = ora(`Converting ${path.basename(input)} to ${options.format}...`).start();

    try {
      const config: WordConfig = {
        preserveFormatting: options.preserveFormatting,
        extractImages: options.images,
        outputFormat: options.format
      };

      const parser = new WordParser(config);
      const result = await parser.parseDocument(input);

      if (!result.success) {
        spinner.fail('Conversion failed');
        if (result.errors?.length) {
          result.errors.forEach(error => console.error(chalk.red(error)));
        }
        process.exit(1);
      }

      if (result.content) {
        const converted = await parser.convertToFormat(result.content, options.format, result.metadata);
        fs.writeFileSync(output, converted, 'utf8');
        spinner.succeed(`Converted to ${chalk.green(output)}`);
      }

    } catch (error) {
      spinner.fail('Conversion failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

program
  .command('info')
  .description('Display Word document information')
  .argument('<input>', 'Input Word document path')
  .option('--detailed', 'Show detailed information')
  .action(async (input: string, options: any) => {
    const spinner = ora('Reading document information...').start();

    try {
      const parser = new WordParser({ includeMetadata: true });
      const result = await parser.parseDocument(input);

      if (!result.success) {
        spinner.fail('Failed to read document');
        if (result.errors?.length) {
          result.errors.forEach(error => console.error(chalk.red(error)));
        }
        process.exit(1);
      }

      spinner.stop();

      console.log(chalk.blue.bold(`\n📄 ${path.basename(input)}`));
      console.log(chalk.gray('─'.repeat(50)));

      if (result.metadata) {
        if (result.metadata.title) {
          console.log(`${chalk.yellow('Title:')} ${result.metadata.title}`);
        }
        if (result.metadata.author) {
          console.log(`${chalk.yellow('Author:')} ${result.metadata.author}`);
        }
        if (result.metadata.subject) {
          console.log(`${chalk.yellow('Subject:')} ${result.metadata.subject}`);
        }
        if (result.metadata.created) {
          console.log(`${chalk.yellow('Created:')} ${result.metadata.created.toLocaleDateString()}`);
        }
        if (result.metadata.modified) {
          console.log(`${chalk.yellow('Modified:')} ${result.metadata.modified.toLocaleDateString()}`);
        }
        if (result.metadata.pages) {
          console.log(`${chalk.yellow('Pages:')} ${result.metadata.pages}`);
        }
        if (result.metadata.words) {
          console.log(`${chalk.yellow('Words:')} ${result.metadata.words}`);
        }
        if (result.metadata.characters) {
          console.log(`${chalk.yellow('Characters:')} ${result.metadata.characters}`);
        }
      }

      if (options.detailed && result.content) {
        const contentLength = result.content.length;
        const paragraphs = result.content.split('\n').filter(p => p.trim()).length;
        
        console.log(chalk.gray('\nContent Analysis:'));
        console.log(`  Content Length: ${contentLength} characters`);
        console.log(`  Paragraphs: ${paragraphs}`);
      }

      if (result.images?.length) {
        console.log(`${chalk.yellow('Images:')} ${result.images.length} found`);
      }

    } catch (error) {
      spinner.fail('Information extraction failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

program
  .command('validate')
  .description('Validate Word document safety and integrity')
  .argument('<input>', 'Input Word document path')
  .action(async (input: string) => {
    const spinner = ora('Validating document...').start();

    try {
      const parser = new WordParser({ safetyChecks: true });
      const result = await parser.parseDocument(input);

      spinner.stop();

      if (result.success) {
        console.log(chalk.green('✅ Document validation passed'));
        
        if (result.warnings?.length) {
          console.log(chalk.yellow('\n⚠️  Warnings:'));
          result.warnings.forEach(warning => console.log(chalk.yellow(`  • ${warning}`)));
        }
      } else {
        console.log(chalk.red('❌ Document validation failed'));
        
        if (result.errors?.length) {
          console.log(chalk.red('\nErrors:'));
          result.errors.forEach(error => console.log(chalk.red(`  • ${error}`)));
        }
        process.exit(1);
      }

    } catch (error) {
      spinner.fail('Validation failed');
      console.error(chalk.red(error instanceof Error ? error.message : 'Unknown error'));
      process.exit(1);
    }
  });

// Global error handler
process.on('uncaughtException', (error) => {
  console.error(chalk.red('Uncaught Exception:'), error.message);
  process.exit(1);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error(chalk.red('Unhandled Rejection at:'), promise, 'reason:', reason);
  process.exit(1);
});

// Parse command line arguments
program.parse();