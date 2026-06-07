/**
 * html2pptx - 核心工具函数
 * 包含颜色转换、坐标换算、字体映射等
 */

// ==================== 常量定义 ====================

// PPT 标准尺寸（单位：英寸）
const PPT_WIDTH_INCH = 13.333;
const PPT_HEIGHT_INCH = 7.5;

// PPT 标准尺寸（单位：pt）
const PPT_WIDTH_PT = 960;
const PPT_HEIGHT_PT = 540;

// H5 标准尺寸（单位：px）
const H5_WIDTH_PX = 1280;
const H5_HEIGHT_PX = 720;

// 默认缩放系数 px → pt
const DEFAULT_SCALE = 0.75;

// 形状最小尺寸（英寸）：避免极小形状在 PPT 中完全不可见，
// 等效约 1 CSS px（96 DPI）。原为 0.1 in（≈ 10px），
// 导致 1-4px 细装饰条被 2-10 倍放大。
const MIN_SHAPE_SIZE_INCH = 0.01;

// 文本框最小尺寸（英寸）：文字框比形状更需要保留兜底下限，
// 避免单字因过小被 PowerPoint 裁切不可见。
const MIN_VISIBLE_TEXT_SIZE_INCH = 0.05;

// 字体白名单
const FONT_WHITELIST = [
  'Microsoft YaHei',
  '微软雅黑',
  'SimHei',
  '黑体',
  'SimSun',
  '宋体',
  'Arial',
  'Times New Roman',
  'Calibri',
  'Courier New',
];

// 默认中文字体
const DEFAULT_CN_FONT = 'Microsoft YaHei';
const DEFAULT_EN_FONT = 'Arial';

// ==================== 颜色转换 ====================

/**
 * 将各种 CSS 颜色格式转换为十六进制
 * @param {string} color - CSS 颜色值
 * @returns {string} 十六进制颜色（如 "FF6600"）
 */
function colorToHex(color) {
  if (!color || color === 'transparent' || color === 'rgba(0, 0, 0, 0)') {
    return null;
  }

  // 已经是十六进制
  if (color.startsWith('#')) {
    return color.replace('#', '').toUpperCase();
  }

  // rgb(r, g, b)
  const rgbMatch = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1]).toString(16).padStart(2, '0');
    const g = parseInt(rgbMatch[2]).toString(16).padStart(2, '0');
    const b = parseInt(rgbMatch[3]).toString(16).padStart(2, '0');
    return `${r}${g}${b}`.toUpperCase();
  }

  // 颜色名称映射
  const colorNames = {
    'red': 'FF0000', 'blue': '0000FF', 'green': '008000',
    'black': '000000', 'white': 'FFFFFF', 'yellow': 'FFFF00',
    'orange': 'FFA500', 'purple': '800080', 'pink': 'FFC0CB',
    'gray': '808080', 'grey': '808080',
    'darkgray': 'A9A9A9', 'lightgray': 'D3D3D3',
    'navy': '000080', 'teal': '008080', 'maroon': '800000',
    'olive': '808000', 'lime': '00FF00', 'aqua': '00FFFF',
    'fuchsia': 'FF00FF', 'silver': 'C0C0C0',
  };

  const lower = color.toLowerCase().trim();
  if (colorNames[lower]) {
    return colorNames[lower];
  }

  // 无法识别的颜色，返回黑色
  return '000000';
}

/**
 * 从 rgba 值中提取透明度 (0-100)
 * @param {string} color - CSS 颜色值
 * @returns {number|null} 透明度百分比，null 表示不透明
 */
function extractTransparency(color) {
  if (!color) return null;
  const rgbaMatch = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)/);
  if (rgbaMatch) {
    const alpha = parseFloat(rgbaMatch[4]);
    return Math.round((1 - alpha) * 100);
  }
  return null;
}

function parseGradient(gradient) {
  if (!gradient || typeof gradient !== 'string') return null;

  const gradientMatch = gradient.match(/linear-gradient\s*\(/i);
  if (!gradientMatch) return null;

  const startIndex = gradientMatch[0].length;
  let content = gradient.slice(startIndex).trim();
  if (content.endsWith(')')) {
    content = content.slice(0, -1).trim();
  }

  let angle = 180;
  const colors = [];
  const parts = [];
  let current = '';
  let parenDepth = 0;

  for (let i = 0; i < content.length; i++) {
    const char = content[i];
    if (char === '(') {
      parenDepth++;
      current += char;
    } else if (char === ')') {
      parenDepth--;
      current += char;
    } else if (char === ',' && parenDepth === 0) {
      parts.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }

  if (current.trim()) {
    parts.push(current.trim());
  }

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const angleMatch = part.match(/(-?\d+(?:\.\d+)?)\s*deg/i);
    if (angleMatch) {
      angle = parseFloat(angleMatch[1]);
      continue;
    }

    if (part.startsWith('to ')) {
      const dir = part.replace('to ', '').trim().toLowerCase();
      if (dir === 'right') angle = 90;
      else if (dir === 'left') angle = 270;
      else if (dir === 'top' || dir === 'up') angle = 0;
      else if (dir === 'bottom' || dir === 'down') angle = 180;
      continue;
    }

    const positionMatch = part.match(/\s+(\d+(?:\.\d+)?)%\s*$/);
    const position = positionMatch ? parseFloat(positionMatch[1]) / 100 : null;

    if (part.toLowerCase().startsWith('transparent')) {
      colors.push({ color: '000000', transparency: 100, position });
      continue;
    }

    const transparentRgbaMatch = part.match(/rgba?\(\s*0\s*,\s*0\s*,\s*0\s*,\s*0\s*\)/i);
    if (transparentRgbaMatch) {
      colors.push({ color: '000000', transparency: 100, position });
      continue;
    }

    const colorMatch = part.match(/(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\))/);
    if (colorMatch) {
      const colorStr = colorMatch[1];
      let transparency = 0;
      const rgbaMatch = colorStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
      if (rgbaMatch) {
        transparency = Math.round((1 - parseFloat(rgbaMatch[4])) * 100);
      }
      colors.push({
        color: colorToHex(colorStr) || 'FFFFFF',
        transparency,
        position,
      });
    }
  }

  if (colors.length < 2) return null;

  const stops = colors.map((c, i) => ({
    color: c.color,
    position: c.position !== null ? c.position : i / (colors.length - 1),
    transparency: c.transparency,
  }));

  return { type: 'linear', angle, stops };
}

/**
 * 处理渐变色背景，取第一个色或平均值
 * @param {string} gradient - CSS 渐变值
 * @returns {string} 十六进制颜色
 */
function gradientToColor(gradient) {
  if (!gradient) return 'FFFFFF';

  // 提取所有颜色值
  const colors = gradient.match(/#[0-9a-fA-F]{3,8}|rgb[a]?\([^)]+\)/g);
  if (colors && colors.length > 0) {
    return colorToHex(colors[0]);
  }

  return 'FFFFFF';
}

// ==================== 坐标换算 ====================

/**
 * px 转 pt
 * @param {number} px - 像素值
 * @param {number} scale - 缩放系数
 * @returns {number} pt 值
 */
function pxToPt(px, scale = DEFAULT_SCALE) {
  return Math.round(px * scale * 100) / 100;
}

/**
 * px 转 英寸（PptxGenJS 使用英寸）
 * @param {number} px - 像素值
 * @param {number} scale - 缩放系数
 * @returns {number} 英寸值
 */
function pxToInch(px, scale = DEFAULT_SCALE) {
  return pxToPt(px, scale) / 72;
}

// ==================== 字体处理 ====================

/**
 * 规范化字体名称，确保在白名单内
 * @param {string} fontFamily - CSS 字体族
 * @returns {string} 规范化后的字体名
 */
function normalizeFont(fontFamily) {
  if (!fontFamily) return DEFAULT_CN_FONT;

  // 拆分字体列表，取第一个在白名单中的
  const fonts = fontFamily.split(',').map(f => f.trim().replace(/['"]/g, ''));

  for (const font of fonts) {
    if (FONT_WHITELIST.some(allowed => font.toLowerCase().includes(allowed.toLowerCase()))) {
      return font;
    }
  }

  // 不在白名单中，返回默认字体
  return DEFAULT_CN_FONT;
}

/**
 * 判断是否是中文字符
 * @param {string} text - 文本
 * @returns {boolean}
 */
function isCJK(text) {
  if (!text) return false;
  return /[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]/.test(text);
}

/**
 * 根据文本内容选择字体
 * @param {string} text - 文本内容
 * @param {string} specifiedFont - 指定的字体
 * @returns {string} 字体名
 */
function resolveFont(text, specifiedFont) {
  const normalized = normalizeFont(specifiedFont);
  if (isCJK(text) && !['Arial', 'Times New Roman', 'Calibri', 'Courier New'].includes(normalized)) {
    return normalized;
  }
  if (isCJK(text)) {
    return DEFAULT_CN_FONT;
  }
  return normalized;
}

// ==================== 样式解析 ====================

/**
 * 解析 CSS 字体大小，返回 pt 值
 * @param {string} fontSize - CSS 字体大小（如 "24px", "16pt", "1.5em"）
 * @returns {number} pt 值
 */
function parseFontSize(fontSize) {
  if (!fontSize) return 12; // 默认 12pt

  const num = parseFloat(fontSize);
  if (isNaN(num)) return 12;

  if (fontSize.includes('px')) {
    return pxToPt(num);
  } else if (fontSize.includes('pt')) {
    return num;
  } else if (fontSize.includes('em')) {
    return num * 12; // 1em ≈ 12pt
  } else if (fontSize.includes('%')) {
    return num * 12 / 100;
  }

  return num;
}

/**
 * 解析 CSS 对齐方式
 * @param {string} align - CSS text-align 值
 * @returns {string} PptxGenJS 对齐值
 */
function parseAlignment(align) {
  switch (align) {
    case 'center': return 'center';
    case 'right': return 'right';
    case 'end': return 'right';
    case 'justify': return 'justify';
    case 'left':
    case 'start':
    default: return 'left';
  }
}

// ==================== 元素分类 ====================

/**
 * 判断元素是否应该生成 PPT 元素
 * @param {object} el - 元素数据
 * @returns {boolean}
 */
function shouldGenerateElement(el) {
  // 跳过 body 和 html 标签
  if (['BODY', 'HTML', 'HEAD', 'STYLE', 'SCRIPT', 'META', 'LINK', 'TITLE'].includes(el.tag)) {
    return false;
  }

  // 跳过没有内容的纯文本节点
  if (!el.text && el.tag !== 'IMG' && el.tag !== 'TABLE' && el.tag !== 'HR') {
    // 检查是否有背景色或边框
    const hasBorder = el.borderWidth > 0 || el.borderTopWidth > 0 || el.borderRightWidth > 0 || 
                      el.borderBottomWidth > 0 || el.borderLeftWidth > 0;
    if (!el.backgroundColor && !hasBorder) {
      return false;
    }
  }

  // 跳过尺寸为 0 的元素
  if (el.width < 1 && el.height < 1) {
    return false;
  }

  return true;
}

/**
 * 判断元素是否是文本类元素
 * @param {string} tag - HTML 标签名
 * @returns {boolean}
 */
function isTextElement(tag) {
  return ['H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'P', 'LI', 'SPAN', 'A', 'B', 'STRONG', 'I', 'EM', 'U'].includes(tag);
}

/**
 * 判断元素是否是容器元素
 * @param {string} tag - HTML 标签名
 * @returns {boolean}
 */
function isContainerElement(tag) {
  return ['DIV', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'MAIN', 'NAV', 'ASIDE'].includes(tag);
}

module.exports = {
  // 常量
  PPT_WIDTH_INCH,
  PPT_HEIGHT_INCH,
  PPT_WIDTH_PT,
  PPT_HEIGHT_PT,
  H5_WIDTH_PX,
  H5_HEIGHT_PX,
  DEFAULT_SCALE,
  MIN_SHAPE_SIZE_INCH,
  MIN_VISIBLE_TEXT_SIZE_INCH,
  FONT_WHITELIST,
  DEFAULT_CN_FONT,
  DEFAULT_EN_FONT,

  // 颜色
  colorToHex,
  extractTransparency,
  gradientToColor,
  parseGradient,

  // 坐标
  pxToPt,
  pxToInch,

  // 字体
  normalizeFont,
  isCJK,
  resolveFont,

  // 样式
  parseFontSize,
  parseAlignment,

  // 元素
  shouldGenerateElement,
  isTextElement,
  isContainerElement,
};
