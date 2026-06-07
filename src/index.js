/**
 * html2pptx - 主入口模块
 * 编排完整的转换流程：HTML 提取 → PPT 生成
 */

const fs = require('fs');
const path = require('path');
const { extractFromHtml } = require('./extract');
const { generatePptx } = require('./generate');
const { DEFAULT_SCALE } = require('./utils');

/**
 * 将一个或多个 HTML 文件转换为 PPT
 * @param {string|string[]} input - HTML 文件路径或路径数组
 * @param {string} outputPath - 输出 .pptx 文件路径
 * @param {object} options - 选项
 * @param {number} options.scale - px→pt 缩放系数，默认 0.75
 * @param {boolean} options.verbose - 是否输出详细日志
 * @returns {Promise<string>} 输出文件路径
 */
async function convert(input, outputPath, options = {}) {
  const { scale = DEFAULT_SCALE, verbose = false } = options;

  // 规范化输入为文件路径数组
  let htmlFiles = [];

  if (Array.isArray(input)) {
    htmlFiles = input;
  } else if (typeof input === 'string') {
    const stat = fs.statSync(input);
    if (stat.isDirectory()) {
      // 目录模式：读取目录下所有 .html 文件，按文件名排序
      htmlFiles = fs.readdirSync(input)
        .filter(f => f.endsWith('.html') || f.endsWith('.htm'))
        .sort()
        .map(f => path.join(input, f));
    } else {
      htmlFiles = [input];
    }
  }

  if (htmlFiles.length === 0) {
    throw new Error('未找到 HTML 文件');
  }

  if (verbose) {
    console.log(`[convert] 输入文件 (${htmlFiles.length} 个):`);
    htmlFiles.forEach(f => console.log(`  - ${f}`));
    console.log(`[convert] 输出文件: ${outputPath}`);
    console.log(`[convert] 缩放系数: ${scale}`);
  }

  // 验证所有文件存在
  for (const file of htmlFiles) {
    if (!fs.existsSync(file)) {
      throw new Error(`文件不存在: ${file}`);
    }
  }

  // 逐个提取 HTML 数据
  const slideDataList = [];
  for (let i = 0; i < htmlFiles.length; i++) {
    const file = htmlFiles[i];
    if (verbose) console.log(`\n[convert] 处理第 ${i + 1}/${htmlFiles.length} 个文件: ${path.basename(file)}`);

    try {
      const slideData = await extractFromHtml(file, {}, verbose);
      slideDataList.push(slideData);
    } catch (err) {
      console.warn(`[convert] 处理文件 ${file} 时出错: ${err.message}`);
      // 跳过出错的文件，继续处理其他文件
    }
  }

  if (slideDataList.length === 0) {
    throw new Error('所有文件处理失败，无法生成 PPT');
  }

  // 生成 PPT
  await generatePptx(slideDataList, outputPath, { scale, verbose });

  if (verbose) {
    console.log(`\n[convert] ✅ 转换完成！`);
    console.log(`[convert] 输入: ${htmlFiles.length} 个 HTML 文件`);
    console.log(`[convert] 输出: ${outputPath}`);
  }

  return outputPath;
}

module.exports = { convert };
