#!/usr/bin/env node
/**
 * html2pptx - CLI 命令行入口
 *
 * 用法：
 *   node src/cli.js input.html -o output.pptx
 *   node src/cli.js slide1.html slide2.html slide3.html -o output.pptx
 *   node src/cli.js ./slides/ -o output.pptx
 *   node src/cli.js input.html -o output.pptx --verbose
 */

const { program } = require('commander');
const path = require('path');
const { convert } = require('./index');

// 版本号
const VERSION = '1.0.0';

program
  .name('html2pptx')
  .description('纯静态 HTML+CSS 转可编辑 PPT 工具')
  .version(VERSION, '-V, --version', '显示版本号')
  .argument('<inputs...>', 'HTML 文件路径或目录路径（支持多个文件，用空格分隔）')
  .option('-o, --output <path>', '输出 .pptx 文件路径', 'output.pptx')
  .option('-s, --scale <number>', 'px→pt 缩放系数', '0.75')
  .option('--verbose', '显示详细日志', false)
  .action(async (inputs, options) => {
    try {
      const outputPath = path.resolve(options.output);
      const scale = parseFloat(options.scale);
      const verbose = options.verbose || false;

      if (verbose) {
        console.log('========================================');
        console.log('  html2pptx v' + VERSION);
        console.log('  纯静态 HTML+CSS → 可编辑 PPT');
        console.log('========================================\n');
      }

      // 如果只有一个输入且是目录，直接传字符串
      // 如果有多个输入，传数组
      const input = inputs.length === 1 ? inputs[0] : inputs;

      await convert(input, outputPath, { scale, verbose });

      console.log(`\n✅ PPT 已生成: ${outputPath}`);
    } catch (err) {
      console.error(`\n❌ 转换失败: ${err.message}`);
      process.exit(1);
    }
  });

// 解析命令行参数
program.parse(process.argv);
