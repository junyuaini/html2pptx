/**
 * html2pptx - 渐变后处理模块
 *
 * 背景：PptxGenJS 3.12.0 没有暴露原生渐变填充 API，ShapeFillProps.type 只允许 'none' | 'solid'。
 * 方案：在 generate.js 生成 PPTX 后，按"slide 索引 + 形状名 Shape K"定位 <p:sp>，
 *      将其中的 <a:solidFill> 替换为符合 OOXML 规范的 <a:gradFill>，
 *      从而获得真正的原生渐变效果（避免分段矩形模拟带来的色阶、斜向不支持、圆角丢失等问题）。
 *
 * 处理顺序：
 *   1. 入参校验
 *   2. 业务校验：文件存在、任务非空
 *   3. 业务处理：解压 → 修改 slideN.xml → 重新打包覆盖
 */

const fs = require('fs');
const JSZip = require('jszip');

const GRADIENT_TASK_KEY = '__html2pptx_gradient_tasks__';

const SUPPORTED_GRADIENT_TYPE = 'linear';

const SHAPE_OPEN_TAG = '<p:sp>';
const SHAPE_CLOSE_TAG = '</p:sp>';

/**
 * 创建一个新的渐变任务收集器
 * @returns {object} { addTask, getTasks, nextShapeName, reset }
 */
function createGradientCollector() {
  // slideIndex (0-based) → 自增序号（仅用于生成 objectName，避免重名）
  const slideSeqCounters = new Map();
  const tasks = [];

  /**
   * 返回 slide 内的下一个唯一序号字符串（如 "0", "1", "2"...）
   * 注意：这里只是序号，不再依赖 PptxGenJS 内部的 "Shape K" 索引；
   * 真正的 OOXML 定位锚点是 objectName（在 generate.js 里拼装）。
   */
  function nextShapeName(slideIndex) {
    if (!Number.isInteger(slideIndex) || slideIndex < 0) {
      throw new Error(`[gradient-postprocess] nextShapeName: slideIndex 非法: ${slideIndex}`);
    }
    const cur = slideSeqCounters.get(slideIndex) || 0;
    slideSeqCounters.set(slideIndex, cur + 1);
    return String(cur);
  }

  function addTask(task) {
    if (!task || typeof task !== 'object') {
      throw new Error('[gradient-postprocess] addTask: task 必须是对象');
    }
    if (!Number.isInteger(task.slideIndex) || task.slideIndex < 0) {
      throw new Error(`[gradient-postprocess] addTask: slideIndex 非法: ${task.slideIndex}`);
    }
    if (typeof task.shapeName !== 'string' || !task.shapeName) {
      throw new Error('[gradient-postprocess] addTask: shapeName 不能为空');
    }
    if (!task.gradient || !Array.isArray(task.gradient.stops) || task.gradient.stops.length < 2) {
      throw new Error('[gradient-postprocess] addTask: gradient.stops 至少需要 2 个');
    }
    tasks.push(task);
  }

  function getTasks() {
    return tasks.slice();
  }

  function reset() {
    slideSeqCounters.clear();
    tasks.length = 0;
  }

  return { addTask, getTasks, nextShapeName, reset };
}

/**
 * 将我们的 gradient.angle（CSS 语义：0deg=向上，顺时针）
 * 转换为 OOXML <a:gradFill> 的 ang 属性（单位为 60000 分之一度，0=向右，顺时针）。
 *
 * CSS → OOXML 的换算：ooxmlDeg = (cssAngle - 90 + 360) % 360
 * 例：CSS 0deg(向上)   → OOXML 270（向上）
 *     CSS 90deg(向右)  → OOXML 0  （向右）
 *     CSS 180deg(向下) → OOXML 90 （向下）
 *     CSS 270deg(向左) → OOXML 180（向左）
 *
 * @param {number} cssAngle
 * @returns {number} OOXML ang 单位值
 */
function cssAngleToOoxmlAng(cssAngle) {
  const normalized = ((Number(cssAngle) % 360) + 360) % 360;
  const ooxmlDeg = (normalized - 90 + 360) % 360;
  return Math.round(ooxmlDeg * 60000);
}

/**
 * 把单个 stop 渲染成 <a:gs> 节点字符串
 * @param {object} stop - { color: 'RRGGBB', position: 0..1, transparency: 0..100 }
 * @returns {string}
 */
function renderGradientStop(stop) {
  const pos = Math.max(0, Math.min(1, Number(stop.position) || 0));
  const posVal = Math.round(pos * 100000);
  const color = String(stop.color || 'FFFFFF').replace('#', '').toUpperCase();
  const alpha = 100 - Math.max(0, Math.min(100, Number(stop.transparency) || 0));
  const alphaVal = Math.round(alpha * 1000);
  const alphaTag = alphaVal < 100000 ? `<a:alpha val="${alphaVal}"/>` : '';
  return `<a:gs pos="${posVal}"><a:srgbClr val="${color}">${alphaTag}</a:srgbClr></a:gs>`;
}

/**
 * 构建 <a:gradFill> XML 片段
 * @param {object} gradient - { type, angle, stops }
 * @returns {string}
 */
function buildGradFillXml(gradient) {
  if (!gradient || gradient.type !== SUPPORTED_GRADIENT_TYPE) {
    return null;
  }
  const stops = [...gradient.stops]
    .map((s, i, arr) => ({
      color: s.color,
      position: s.position !== null && s.position !== undefined
        ? s.position
        : i / Math.max(arr.length - 1, 1),
      transparency: s.transparency || 0,
    }))
    .sort((a, b) => a.position - b.position);

  const ang = cssAngleToOoxmlAng(gradient.angle);
  const gsList = stops.map(renderGradientStop).join('');
  return `<a:gradFill rotWithShape="1"><a:gsLst>${gsList}</a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
}

/**
 * 在一段 slide xml 中按 shapeName 定位 <p:sp> 区间
 * @param {string} xml
 * @param {string} shapeName
 * @returns {{start:number,end:number}|null}
 */
function locateShapeRange(xml, shapeName) {
  // 找到 cNvPr 中 name="Shape K"
  const needle = `name="${shapeName}"`;
  const namePos = xml.indexOf(needle);
  if (namePos < 0) return null;

  // 向前找最近的 <p:sp>
  const start = xml.lastIndexOf(SHAPE_OPEN_TAG, namePos);
  if (start < 0) return null;

  // 向后找最近的 </p:sp>
  const closeIdx = xml.indexOf(SHAPE_CLOSE_TAG, namePos);
  if (closeIdx < 0) return null;
  const end = closeIdx + SHAPE_CLOSE_TAG.length;

  return { start, end };
}

/**
 * 在形状 XML 片段中，把第一个 <a:solidFill>...</a:solidFill> 替换为 gradFill
 * @param {string} spXml
 * @param {string} gradFillXml
 * @returns {string|null} 替换后的 xml；找不到 solidFill 则返回 null
 */
function replaceSolidFillWithGradient(spXml, gradFillXml) {
  // 只替换 spPr 内的第一个 solidFill（避免误改 line/text 内的 solidFill）
  const spPrStart = spXml.indexOf('<p:spPr');
  const spPrEnd = spXml.indexOf('</p:spPr>');
  if (spPrStart < 0 || spPrEnd < 0 || spPrEnd <= spPrStart) return null;

  const before = spXml.slice(0, spPrStart);
  const spPrBlock = spXml.slice(spPrStart, spPrEnd);
  const after = spXml.slice(spPrEnd);

  const solidStart = spPrBlock.indexOf('<a:solidFill>');
  if (solidStart < 0) return null;
  const solidEnd = spPrBlock.indexOf('</a:solidFill>', solidStart);
  if (solidEnd < 0) return null;
  const newSpPrBlock =
    spPrBlock.slice(0, solidStart) +
    gradFillXml +
    spPrBlock.slice(solidEnd + '</a:solidFill>'.length);

  return before + newSpPrBlock + after;
}

/**
 * 处理 slide 背景（<p:bg>）的渐变替换
 * 背景没有 cNvPr 锚点，只有一个 bgPr 节点，直接替换它的 solidFill
 * @param {string} xml
 * @param {string} gradFillXml
 * @returns {string|null}
 */
function replaceSlideBackgroundFill(xml, gradFillXml) {
  const bgStart = xml.indexOf('<p:bg>');
  const bgEnd = xml.indexOf('</p:bg>');
  if (bgStart < 0 || bgEnd < 0 || bgEnd <= bgStart) return null;

  const bgBlock = xml.slice(bgStart, bgEnd);
  const solidStart = bgBlock.indexOf('<a:solidFill>');
  if (solidStart < 0) return null;
  const solidEnd = bgBlock.indexOf('</a:solidFill>', solidStart);
  if (solidEnd < 0) return null;

  const newBgBlock =
    bgBlock.slice(0, solidStart) +
    gradFillXml +
    bgBlock.slice(solidEnd + '</a:solidFill>'.length);

  return xml.slice(0, bgStart) + newBgBlock + xml.slice(bgEnd);
}

/**
 * 应用所有渐变任务到 pptx 文件
 * @param {string} pptxPath - 已生成的 pptx 文件路径（会被覆盖写回）
 * @param {Array} tasks - 渐变任务列表
 * @param {object} [options]
 * @param {boolean} [options.verbose]
 * @returns {Promise<{applied:number, skipped:number}>}
 */
async function applyGradients(pptxPath, tasks, options = {}) {
  const { verbose = false } = options;

  // 1. 入参校验
  if (typeof pptxPath !== 'string' || !pptxPath) {
    throw new Error('[gradient-postprocess] applyGradients: pptxPath 不能为空');
  }
  if (!Array.isArray(tasks)) {
    throw new Error('[gradient-postprocess] applyGradients: tasks 必须是数组');
  }

  // 2. 业务校验
  if (tasks.length === 0) {
    if (verbose) console.log('[gradient-postprocess] 无渐变任务，跳过后处理');
    return { applied: 0, skipped: 0 };
  }
  if (!fs.existsSync(pptxPath)) {
    throw new Error(`[gradient-postprocess] 文件不存在: ${pptxPath}`);
  }

  // 3. 业务处理
  const buf = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(buf);

  // 按 slideIndex 分桶
  const tasksBySlide = new Map();
  for (const t of tasks) {
    if (!tasksBySlide.has(t.slideIndex)) {
      tasksBySlide.set(t.slideIndex, []);
    }
    tasksBySlide.get(t.slideIndex).push(t);
  }

  let applied = 0;
  let skipped = 0;

  for (const [slideIndex, slideTasks] of tasksBySlide.entries()) {
    const slideFile = `ppt/slides/slide${slideIndex + 1}.xml`;
    const entry = zip.file(slideFile);
    if (!entry) {
      if (verbose) console.warn(`[gradient-postprocess] 找不到 ${slideFile}，跳过 ${slideTasks.length} 个任务`);
      skipped += slideTasks.length;
      continue;
    }

    let xml = await entry.async('string');

    for (const task of slideTasks) {
      const gradFillXml = buildGradFillXml(task.gradient);
      if (!gradFillXml) {
        if (verbose) console.warn(`[gradient-postprocess] 跳过任务（gradient 不支持）: slide=${slideIndex} shape=${task.shapeName}`);
        skipped++;
        continue;
      }

      // 背景任务：shapeName 为特殊值 '__bg__'
      if (task.shapeName === '__bg__') {
        const next = replaceSlideBackgroundFill(xml, gradFillXml);
        if (!next) {
          if (verbose) console.warn(`[gradient-postprocess] 背景替换失败: slide=${slideIndex}`);
          skipped++;
          continue;
        }
        xml = next;
        applied++;
        continue;
      }

      const range = locateShapeRange(xml, task.shapeName);
      if (!range) {
        if (verbose) console.warn(`[gradient-postprocess] 找不到形状: slide=${slideIndex} name=${task.shapeName}`);
        skipped++;
        continue;
      }
      const spXml = xml.slice(range.start, range.end);
      const newSpXml = replaceSolidFillWithGradient(spXml, gradFillXml);
      if (!newSpXml) {
        if (verbose) console.warn(`[gradient-postprocess] 形状内未找到 solidFill: slide=${slideIndex} name=${task.shapeName}`);
        skipped++;
        continue;
      }
      xml = xml.slice(0, range.start) + newSpXml + xml.slice(range.end);
      applied++;
    }

    zip.file(slideFile, xml);
  }

  const outBuf = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  fs.writeFileSync(pptxPath, outBuf);

  if (verbose) {
    console.log(`[gradient-postprocess] 已应用 ${applied} 个渐变，跳过 ${skipped} 个`);
  }
  return { applied, skipped };
}

module.exports = {
  GRADIENT_TASK_KEY,
  createGradientCollector,
  applyGradients,
  buildGradFillXml,
  cssAngleToOoxmlAng,
};
