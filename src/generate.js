/**
 * html2pptx - PPT 生成模块
 * 将提取的元素数据转换为 PPT 幻灯片
 */

const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const {
  colorToHex,
  extractTransparency,
  gradientToColor,
  parseGradient,
  pxToInch,
  parseFontSize,
  parseAlignment,
  normalizeFont,
  resolveFont,
  isTextElement,
  isContainerElement,
  shouldGenerateElement,
  PPT_WIDTH_INCH,
  PPT_HEIGHT_INCH,
  DEFAULT_SCALE,
  DEFAULT_CN_FONT,
  MIN_SHAPE_SIZE_INCH,
  MIN_VISIBLE_TEXT_SIZE_INCH,
} = require('./utils');
const { createGradientCollector, applyGradients } = require('./gradient-postprocess');

// 渐变占位形状的 objectName 前缀，后处理模块据此精确定位
const GRADIENT_OBJECT_NAME_PREFIX = '__html2pptx_grad_';

/**
 * 用纯色占位形状 + 后处理任务的方式，添加一个"原生渐变填充"形状。
 * 由于 PptxGenJS 3.12.0 没有暴露 gradFill API，我们：
 *   1. 用占位色（取渐变首色）先生成一个标准 addShape，确保形状结构（圆角、边框、ext 等）完整；
 *   2. 给它指定一个全局唯一的 objectName（PptxGenJS 会直接写入 <p:cNvPr name="..."/>）；
 *   3. 把 { slideIndex, shapeName, gradient } 推入收集器，generatePptx 写文件后做后处理。
 *
 * @param {object} ctx - 渲染上下文 { slide, collector, slideIndex }
 * @param {string} shapeType - PptxGenJS shape 类型，如 'rect' / 'roundRect' / 'ellipse'
 * @param {object} shapeOptions - 形状属性（不要传 fill；本函数会注入占位 fill 和 objectName）
 * @param {object} gradient - parseGradient 返回的对象 { type, angle, stops }
 */
function addNativeGradientShape(ctx, shapeType, shapeOptions, gradient) {
  if (!ctx || !ctx.slide || !ctx.collector) {
    throw new Error('[generate] addNativeGradientShape: ctx 缺失 slide 或 collector');
  }
  if (!gradient || !Array.isArray(gradient.stops) || gradient.stops.length < 2) {
    throw new Error('[generate] addNativeGradientShape: gradient 非法');
  }

  const slideIndex = ctx.slideIndex;
  const shapeName = ctx.collector.nextShapeName(slideIndex);
  const uniqueObjectName = `${GRADIENT_OBJECT_NAME_PREFIX}${slideIndex}_${shapeName.replace(/\s+/g, '_')}__`;

  // 占位色：首个 stop 的颜色，确保即使后处理失败也有可见效果（降级）
  const placeholderColor = (gradient.stops[0] && gradient.stops[0].color) || 'FFFFFF';
  const finalOptions = {
    ...shapeOptions,
    fill: { type: 'solid', color: placeholderColor },
    objectName: uniqueObjectName,
  };

  ctx.slide.addShape(shapeType, finalOptions);

  ctx.collector.addTask({
    slideIndex,
    shapeName: uniqueObjectName,
    gradient,
  });
}

/**
 * 从提取数据生成一页 PPT 幻灯片
 * @param {PptxGenJS} pptx - PptxGenJS 实例
 * @param {object} slideData - 提取的页面数据 { elements, tables, pageWidth, pageHeight }
 * @param {object} options - 选项 { scale, verbose, collector, slideIndex }
 *   - collector: 渐变任务收集器（来自 createGradientCollector），可选；
 *                若提供且对应 slideIndex 也提供，则使用"原生渐变"（后处理替换 OOXML）；
 *                否则仅保留首色纯色背景作为兜底。
 *   - slideIndex: 当前 slide 在整个 pptx 中的 0-based 索引
 */
function generateSlide(pptx, slideData, options = {}) {
  const { scale = DEFAULT_SCALE, verbose = false, collector = null, slideIndex = null } = options;
  const { elements, tables, pageWidth, pageHeight, bodyBackgroundColor } = slideData;

  // 创建幻灯片
  const slide = pptx.addSlide();

  // 渲染上下文：是否启用"原生渐变"路径
  const useNativeGradient = collector && Number.isInteger(slideIndex) && slideIndex >= 0;
  const gradCtx = useNativeGradient ? { slide, collector, slideIndex } : null;
  // 挂载到 slide 上，便于 generateContainer / generateCircle 等子函数取用，
  // 避免一次性修改大量函数签名引入回归风险
  slide.__html2pptxGradCtx = gradCtx;

  // 设置幻灯片背景色
  if (bodyBackgroundColor && bodyBackgroundColor !== 'transparent' && bodyBackgroundColor !== 'rgba(0, 0, 0, 0)') {
    if (bodyBackgroundColor.includes('gradient')) {
      const bgGradient = parseGradient(bodyBackgroundColor);
      if (bgGradient) {
        if (useNativeGradient) {
          // 原生渐变路径：先设占位纯色背景，再登记一个 __bg__ 任务，由后处理替换为 <a:gradFill>
          slide.background = { type: 'solid', color: gradientToColor(bodyBackgroundColor) };
          collector.addTask({
            slideIndex,
            shapeName: '__bg__',
            gradient: bgGradient,
          });
        } else {
          // 兜底（理论上不会走到，generatePptx 默认始终传入 collector）：
          // 仅设首色纯色背景，避免完全失色
          slide.background = { type: 'solid', color: gradientToColor(bodyBackgroundColor) };
        }
      }
    } else if (/url\(["']?data:image\//.test(bodyBackgroundColor)) {
      // 自定义修改：data: URL 背景图（如 .slide 的 base64 嵌入图），用 addImage 全幅平铺
      // 抽取 url("...") 里的 data 部分
      // 分组：m[1]=完整 dataURL; m[2]=mime(jpeg/png); m[3]=base64 内容
      const m = bodyBackgroundColor.match(/url\(["']?(data:image\/([^;)"']+);base64,([^"')]+))["']?\)/);
      const mime = m ? m[2] : null;  // 修：原用 m[1] 错（拿到完整 dataURL 当 mime）
      const base64 = m ? m[3] : null;
      if (base64) {
        try {
          // 落盘临时文件，用 path 方式 addImage（更稳）
          const ext = (mime || 'png').replace('jpeg', 'jpg');
          const tmpDir = path.join(require('os').tmpdir(), 'html2pptx-bg');
          fs.mkdirSync(tmpDir, { recursive: true });
          const tmpFile = path.join(tmpDir, `bg-${Date.now()}-${Math.random().toString(36).slice(2)}.${ext}`);
          fs.writeFileSync(tmpFile, Buffer.from(base64, 'base64'));
          slide.addImage({
            path: tmpFile,
            x: 0,
            y: 0,
            w: PPT_WIDTH_INCH,
            h: PPT_HEIGHT_INCH,
          });
        } catch (err) {
          console.warn('[generate] body bg addImage failed:', err.message);
        }
      }
    } else {
      const bgColor = colorToHex(bodyBackgroundColor);
      if (bgColor) {
        const bgOptions = { type: 'solid', color: bgColor };
        const transparency = extractTransparency(bodyBackgroundColor);
        if (transparency !== null && transparency > 0) {
          bgOptions.transparency = transparency;
        }
        slide.background = bgOptions;
      }
    }
  }

  // 幻灯片尺寸由 pptx.layout = 'LAYOUT_WIDE' 统一控制（13.333 x 7.5 英寸）

  // 收集已处理的文本元素位置，避免重复生成
  const processedAreas = [];

  // 判断一个元素是否应该被跳过
  // 1. 完全在已处理区域内 → 跳过（避免重复）
  // 2. 与表格区域重叠 → 跳过（避免覆盖表格）
  // 3. 背景容器例外：即使包含已处理区域，也不跳过（需要生成背景）
  // 4. 圆形元素例外：圆形元素即使在已处理区域内，也应该生成（因为它有独立的视觉样式）
  function shouldSkip(el) {
    const elX = el.x;
    const elY = el.y;
    const elW = el.width;
    const elH = el.height;

    // 检查是否是背景容器（有背景色且无文字的容器）
    const isBackgroundContainer = isContainerElement(el.tag) && el.backgroundColor && (!el.text || el.text.length === 0);

    // 圆形元素总是生成（即使在已处理区域内）
    // 因为圆形通常有独立的背景或边框样式，不应该被跳过
    const isCircleElement = el.isCircle && el.backgroundColor;

    for (const area of processedAreas) {
      // 完全包含：当前元素完全在已处理区域内
      if (elX >= area.x && elY >= area.y &&
          elX + elW <= area.x + area.w &&
          elY + elH <= area.y + area.h) {
        // 圆形元素例外：即使在已处理区域内，也不跳过
        if (isCircleElement) {
          continue; // 不跳过，继续检查其他区域
        }
        // 背景容器例外：即使包含已处理区域，也不跳过（需要生成背景）
        if (isBackgroundContainer) {
          continue; // 不跳过，继续检查其他区域
        }
        return true;
      }
      // 表格区域特殊处理：任何与表格重叠的元素都跳过
      if (area.isTable &&
          elX < area.x + area.w && elX + elW > area.x &&
          elY < area.y + area.h && elY + elH > area.y) {
        return true;
      }
    }
    return false;
  }

  // 处理表格
  if (tables && tables.length > 0) {
    tables.forEach(table => {
      if (verbose) console.log(`[generate] 处理表格: ${table.rows.length} 行`);
      generateTable(slide, table, scale);
      // 标记表格区域为已处理（特殊标记，任何重叠元素都跳过）
      processedAreas.push({ x: table.x, y: table.y, w: table.width, h: table.height, isTable: true });
    });
  }

  // 处理元素（按层次排序：背景容器先，其他元素后）
  // 分层策略：
  // 1. 背景层：有背景色的容器元素（如 .title-box, .chart-bar-box 等）
  const pseudoElements = slideData.pseudoElements || [];
  const lowerPseudoElements = pseudoElements.filter(pseudo => {
    const parentClassName = (pseudo.parentClassName || '').toString();
    const zIndex = parseInt(pseudo.zIndex, 10);
    const parentZIndex = parseInt(pseudo.parentZIndex, 10);
    // === 方案B：扩大下层伪元素判定 ===
    // 1) 父类为 flow-line（串联线）：明确下层
    // 2) 伪元素显式 z-index <= 父元素 z-index：下层或同层（z-index:0 也算显式声明）
    // 3) 伪元素 z-index < 0：负 z-index 一定在所有元素之下
    const hasFlowLineClass = parentClassName.split(/\s+/).includes('flow-line');
    const hasLowerZ = Number.isFinite(zIndex) && (
      (Number.isFinite(parentZIndex) && zIndex <= parentZIndex) ||
      (!Number.isFinite(parentZIndex) && zIndex <= 0)
    );
    return hasFlowLineClass || hasLowerZ;
  });
  const upperPseudoElements = pseudoElements.filter(pseudo => !lowerPseudoElements.includes(pseudo));

  let lowerPseudoElementsGenerated = false;

  // 1. 背景层：容器背景、形状
  // 2. 内容层：文本、图片等
  // 3. SVG 层
  // 4. 装饰层：需要置顶的伪元素
  // BUG 修复：装饰元素（className 以 deco- 开头）排在页面背景之后、内容之前
  // 避免装饰元素被 .slide 主背景覆盖（如 .deco-circle）
  // === 方案B：在排序中引入 CSS z-index 作为权重 ===
  // 优先级从低到高：
  //   1) z-index 显式小于 0 的元素（最底层，负 z-index）
  //   2) 装饰元素（className 以 deco- 开头）—— 排在页面背景之后
  //   3) 背景容器（有背景色且无文字的容器）
  //   4) 普通内容元素（默认层）
  //   5) z-index 显式大于 0 的元素（最顶层，正 z-index）
  // 同层按 y 坐标、再按 x 坐标排序
  const isDecoElement = (el) => (el.className || '').toString().split(/\s+/).some(c => c.startsWith('deco-'));
  const getEffectiveZIndex = (el) => (el && Number.isFinite(el.zIndex)) ? el.zIndex : 0;
  const sortedElements = [...elements].sort((a, b) => {
    // 1) 负 z-index 排最前
    const aZ = getEffectiveZIndex(a);
    const bZ = getEffectiveZIndex(b);
    if (aZ < 0 && bZ >= 0) return -1;
    if (aZ >= 0 && bZ < 0) return 1;
    // 2) 装饰元素优先级最低（页面背景之后，内容元素之后）
    const aIsDeco = isDecoElement(a);
    const bIsDeco = isDecoElement(b);
    if (aIsDeco && !bIsDeco) return 1;
    if (!aIsDeco && bIsDeco) return -1;
    // 3) 背景容器优先（有背景色且无文字的容器）
    const aIsBg = isContainerElement(a.tag) && a.backgroundColor && (!a.text || a.text.length === 0);
    const bIsBg = isContainerElement(b.tag) && b.backgroundColor && (!b.text || b.text.length === 0);
    if (aIsBg && !bIsBg) return -1;
    if (!aIsBg && bIsBg) return 1;
    // 4) 正 z-index 排最后（最上层）
    if (aZ > 0 && bZ === 0) return 1;
    if (aZ === 0 && bZ > 0) return -1;
    // 同层按 y 坐标排序
    return a.y - b.y || a.x - b.x;
  });

  sortedElements.forEach((el, idx) => {
    const className = (el.className || '').toString();
    if (!lowerPseudoElementsGenerated && className.split(/\s+/).includes('step')) {
      lowerPseudoElements.forEach(pseudo => {
        if (verbose) console.log(`[generate] 处理底层伪元素: ${pseudo.text || '(shape)'}`);
        generatePseudoElement(slide, pseudo, scale);
      });
      lowerPseudoElementsGenerated = true;
    }

    // 跳过不应生成的元素
    if (!shouldGenerateElement(el)) return;

    // 跳过重复区域内的元素
    if (shouldSkip(el)) {
      return;
    }

    // 跳过表格内的 td/th（已通过表格处理）
    if (['TD', 'TH', 'TR', 'TBODY', 'THEAD', 'TFOOT', 'CAPTION'].includes(el.tag)) return;

    // 跳过列表容器（只处理 li）
    if (['UL', 'OL'].includes(el.tag)) return;

    try {
      if (el.tag === 'IMG') {
        generateImagePlaceholder(slide, el, scale);
      } else if (el.tag === 'HR') {
        generateLine(slide, el, scale);
      } else if (isTextElement(el.tag)) {
        // 检查是否有背景色或边框（如 .tag 元素）
        const hasBackground = el.backgroundColor && el.backgroundColor !== 'transparent' && el.backgroundColor !== 'rgba(0, 0, 0, 0)';
        const hasAnyBorder = el.borderWidth > 0 || 
                            el.borderTopWidth > 0 || 
                            el.borderRightWidth > 0 || 
                            el.borderBottomWidth > 0 || 
                            el.borderLeftWidth > 0;
        
        // 如果文本元素有背景色或边框，先生成背景
        if (hasBackground || hasAnyBorder) {
          if (el.isCircle) {
            generateCircle(slide, el, scale);
            processedAreas.push({ x: el.x, y: el.y, w: el.width, h: el.height });
          } else {
            generateContainer(slide, el, scale);
          }
        }
        
        // 检查是否有对应的伪元素装饰条，如果有则文本向右偏移
        // 只对装饰条（isDecorLine=true）应用偏移，不影响标题等普通文本
        // 检查：伪元素在文本元素范围内（x >= el.x 且在 el.x + 20px 内），且 y 坐标在文本范围内
        const pseudo = slideData.pseudoElements ? slideData.pseudoElements.find(p =>
          p.isDecorLine && p.x >= el.x && p.x < el.x + 20 && p.y >= el.y && p.y < el.y + el.height
        ) : null;
        // === BUG 修复：跳过已被父元素合并到 rich text runs 的子 SPAN 文本生成 ===
        // 仅生成 SPAN 的 background 矩形（高亮底色），不生成独立文本框（避免重复）
        if (el.tag === 'SPAN' && el._mergedIntoParentRuns && el.text) {
          // 不调用 generateTextElement，跳过独立文本生成
          processedAreas.push({ x: el.x, y: el.y, w: el.width, h: el.height });
          return;
        }

        if (pseudo) {
          // 文本向右偏移：伪元素宽度 + 伪元素margin-right（如果有）
          // ::before装饰条的offset = width + marginRight，确保装饰条和文本之间有正确的间距
          const offsetX = pseudo.width + (pseudo.marginRight || 9);
          const adjustedEl = { ...el, x: el.x + offsetX, width: el.width - offsetX };
          generateTextElement(slide, adjustedEl, scale);
        } else {
          generateTextElement(slide, el, scale);
        }
        processedAreas.push({ x: el.x, y: el.y, w: el.width, h: el.height });
      } else if (isContainerElement(el.tag)) {
        // 容器元素：有背景就生成形状，有边框也要生成（用于坐标轴等）
        const hasBackground = el.backgroundColor && el.backgroundColor !== 'transparent' && el.backgroundColor !== 'rgba(0, 0, 0, 0)';
        const hasAnyBorder = el.borderWidth > 0 || 
                            el.borderTopWidth > 0 || 
                            el.borderRightWidth > 0 || 
                            el.borderBottomWidth > 0 || 
                            el.borderLeftWidth > 0;
        
        if (hasBackground || hasAnyBorder) {
          if (el.isCircle) {
            generateCircle(slide, el, scale);
          } else {
            generateContainer(slide, el, scale);
          }
        }
        if (el.text && el.text.length > 0) {
          // 检查是否有对应的伪元素装饰条，如果有则文本向右偏移
          // 只对装饰条（isDecorLine=true）应用偏移，不影响标题等普通文本
          // 检查：伪元素在文本元素范围内（x >= el.x 且在 el.x + 20px 内），且 y 坐标在文本范围内
          const pseudo = slideData.pseudoElements ? slideData.pseudoElements.find(p =>
            p.isDecorLine && p.x >= el.x && p.x < el.x + 20 && p.y >= el.y && p.y < el.y + el.height
          ) : null;
          if (pseudo) {
            // 文本向右偏移：伪元素宽度 + 伪元素margin-right（如果有）
            // ::before装饰条的offset = width + marginRight，确保装饰条和文本之间有正确的间距
            const offsetX = pseudo.width + (pseudo.marginRight || 9);
            const adjustedEl = { ...el, x: el.x + offsetX, width: el.width - offsetX };
            generateTextElement(slide, adjustedEl, scale);
          } else {
            generateTextElement(slide, el, scale);
          }
          processedAreas.push({ x: el.x, y: el.y, w: el.width, h: el.height });
        }
      }
    } catch (err) {
      if (verbose) console.warn(`[generate] 处理元素 ${el.tag} 时出错: ${err.message}`);
    }
  });

  if (!lowerPseudoElementsGenerated && lowerPseudoElements.length > 0) {
    lowerPseudoElements.forEach(pseudo => {
      if (verbose) console.log(`[generate] 处理底层伪元素: ${pseudo.text || '(shape)'}`);
      generatePseudoElement(slide, pseudo, scale);
    });
  }

  // 处理 SVG（转为图片嵌入）— 在内容层之上
  if (slideData.svgs && slideData.svgs.length > 0) {
    slideData.svgs.forEach(svg => {
      if (verbose) console.log(`[generate] 处理 SVG: ${svg.width}x${svg.height}`);
      generateSvgImage(slide, svg, scale);
    });
  }

  // 处理伪元素 — 在最上层
  if (upperPseudoElements.length > 0) {
    upperPseudoElements.forEach(pseudo => {
      if (verbose) console.log(`[generate] 处理伪元素: ${pseudo.text || '(shape)'}`);
      generatePseudoElement(slide, pseudo, scale);
    });
  }

  return slide;
}

/**
 * 生成文本元素
 */
function generateTextElement(slide, el, scale) {
  let x = pxToInch(el.x, scale);
  const y = pxToInch(el.y, scale);
  let w = pxToInch(el.width, scale);
  const h = pxToInch(el.height, scale);

  if (el.textOffsetLeft > 0) {
    const textOffsetLeft = pxToInch(el.textOffsetLeft, scale);
    x += textOffsetLeft;
    w = Math.max(0.01, w - textOffsetLeft);
  }

  const color = colorToHex(el.color);
  const fontSize = parseFontSize(el.fontSize);
  const font = resolveFont(el.text || '', el.fontFamily);
  
  // 判断是否有背景色或边框
  const hasBackground = el.backgroundColor && el.backgroundColor !== 'transparent' && el.backgroundColor !== 'rgba(0, 0, 0, 0)';
  const hasAnyBorder = el.borderWidth > 0 || 
                       el.borderTopWidth > 0 || 
                       el.borderRightWidth > 0 || 
                       el.borderBottomWidth > 0 || 
                       el.borderLeftWidth > 0;
  
  let align = parseAlignment(el.textAlign);
  if (el.isCircle) {
    align = 'center';
  }
  if (el.display === 'flex' && el.justifyContent === 'center') {
    align = 'center';
  }

  const bold = el.fontWeight === 'bold' || parseInt(el.fontWeight) >= 600;
  const italic = el.fontStyle === 'italic';
  const underline = el.textDecoration && el.textDecoration.includes('underline');

  // 判断是否是列表项
  const isListItem = el.tag === 'LI';

  const text = el.text || '';

  // 获取 padding 值（像素）
  const paddingTop = el.padding ? el.padding.top || 0 : 0;
  const paddingBottom = el.padding ? el.padding.bottom || 0 : 0;
  
  // 计算内容高度（去除 padding 的影响）
  const contentHeightPx = el.height - (paddingTop + paddingBottom);
  const contentHeightInch = pxToInch(contentHeightPx, scale);
  
  // 使用实际的 line-height（如果存在）或默认值
  const lineHeightValue = el.lineHeight ? parseFloat(el.lineHeight) : (fontSize * 1.5);
  const lineHeightInch = (isNaN(lineHeightValue) ? fontSize * 1.5 : lineHeightValue) / 72;
  
  // 改进的单行文本判断：考虑 padding，使用实际 line-height
  const isSingleLineInH5 = contentHeightInch <= lineHeightInch * 1.5;

  // 检查 CSS 样式指定的对齐方式
  let cssSpecifiedMiddleAlign = false;
  
  // 1. flex 容器的 align-items: center
  if (el.display === 'flex' && el.alignItems === 'center') {
    cssSpecifiedMiddleAlign = true;
  }
  
  // 2. vertical-align: middle（对 inline/inline-block 元素）
  if (el.verticalAlign === 'middle') {
    cssSpecifiedMiddleAlign = true;
  }
  
  // 3. 表格单元格通常垂直居中
  if (el.tag === 'TD' || el.tag === 'TH') {
    cssSpecifiedMiddleAlign = true;
  }
  
  // 综合判断垂直对齐方式
  // 注意：多行文本即使有背景色也使用顶端对齐（除非CSS明确指定居中）
  // 修改：强制所有文字垂直居中
  const shouldMiddleAlign = true; // 强制所有文字垂直居中
  
  const useMiddleAlign = shouldMiddleAlign;

  const hasRichTextRuns = el.runs && el.runs.length > 1;
  const isBlockText = el.tag === 'DIV' || el.tag === 'P' || el.tag === 'LI';
  const isNoWrap = el.whiteSpace === 'nowrap' || el.whiteSpace === 'pre';
  const shouldWrap = isNoWrap ? false : (isBlockText || hasRichTextRuns || !isSingleLineInH5);

  const MULTI_LINE_BUFFER_PT = 2;
  const bufferInch = shouldWrap ? MULTI_LINE_BUFFER_PT / 72 : 0;
  const textW = w + bufferInch;

  console.log(`[GENERATE-TEXT-DEBUG] tag=${el.tag}, text=${el.text.substring(0, 30)}${el.text.length > 30 ? '...' : ''}, `
    + `contentHeight=${contentHeightInch.toFixed(4)}in, lineHeight=${lineHeightInch.toFixed(4)}in, `
    + `isSingleLine=${isSingleLineInH5}, shouldWrap=${shouldWrap}, hasRichTextRuns=${hasRichTextRuns}, cssAlign=${cssSpecifiedMiddleAlign}, `
    + `display=${el.display}, whiteSpace=${el.whiteSpace}, alignItems=${el.alignItems}, verticalAlign=${el.verticalAlign}, `
    + `useMiddleAlign=${useMiddleAlign}, valign=${useMiddleAlign ? 'middle' : 'top'}`);
  console.log(`[GENERATE-TEXT] tag=${el.tag}, text="${el.text}", x=${x}, y=${y}, isCircle=${el.isCircle}, hasBackground=${hasBackground}, hasAnyBorder=${hasAnyBorder}, useMiddleAlign=${useMiddleAlign}, textAlign=${el.textAlign}, valign=${useMiddleAlign ? 'middle' : 'top'}`);
  // 垂直居中：恢复正确的文字位置计算，修复边框问题
  // 用户反馈：修改后文字位置都往下来了，之前的位置是正确的
  // 恢复修改前的正确计算：textY = y + (h - lineHeightInch) / 2
  // 其中 lineHeightInch = fontSize / 72（单行文字高度）
  let textY = y;
  let textH = h;
  
  if (useMiddleAlign) {
    // 恢复修改前的正确计算：使用 fontSize / 72 作为单行文字高度
    const lineHeightInch = fontSize / 72; // 单行文字高度（inch）
    
    if (el.isCircle || hasBackground || hasAnyBorder) {
      // 有背景/边框/圆形元素的问题修复
      // 修改前：使用方式A（有边距 + valign），导致双重偏移
      // 修改后：手动计算居中位置，文本框高度 = 原始高度，使用 valign
      textY = y;
      textH = h;  // 文本框高度 = 原始高度
    } else {
      // 普通元素：恢复修改前的正确计算
      textY = y;
      textH = h;
    }
  }
  
  console.log(`[GENERATE-TEXT-POS] tag=${el.tag}, text=${el.text}, textY=${textY}, textH=${textH}`);

  // 决定是否真正使用 valign
  // 修改：所有文字都使用 valign 垂直居中
  const actuallyUseValign = useMiddleAlign;
  
  // 构建 text options
  const textOptions = {
    x,
    y: textY,
    w: Math.max(textW, 0.1),
    h: Math.max(textH, 0.1),
    fontSize,
    fontFace: font,
    color: color || '000000',
    align,
    valign: 'middle', // 强制所有文字垂直居中
    bold,
    italic,
    underline,
    wrap: shouldWrap,
    margin: [
      (el.padding.left || 0) * 0.75,
      (el.padding.right || 0) * 0.75,
      0,
      0,
    ],
  };

  // 处理富文本 runs
  if (el.runs && el.runs.length > 1) {
    // 多个 run，使用 richText
    const richTextRuns = el.runs.map(run => {
      const runColor = colorToHex(run.color);
      const runFontSize = parseFontSize(run.fontSize);
      const runFont = resolveFont(run.text || '', run.fontFamily);

      return {
        text: run.text,
        options: {
          fontSize: runFontSize,
          fontFace: runFont,
          color: runColor || '000000',
          bold: run.bold || false,
          italic: run.italic || false,
          underline: run.underline || false,
          breakLine: run.breakLine || false,
        },
      };
    });

    slide.addText(richTextRuns, textOptions);
  } else {
    // 单一文本
    let displayText = text;

    // 列表项加项目符号
    if (isListItem) {
      displayText = '• ' + displayText;
    }

    if (displayText) {
        slide.addText(displayText, textOptions);
    }
  }
}

/**
 * 生成图片占位框
 * 增强（自定义修改）：当 imgSrc 存在且能读到本地文件时，用 addImage 嵌入图片
 * 否则回退到占位框
 */
function generateImagePlaceholder(slide, el, scale) {
  const x = pxToInch(el.x, scale);
  const y = pxToInch(el.y, scale);
  const w = pxToInch(el.width, scale);
  const h = pxToInch(el.height, scale);

  // 尝试读取图片二进制
  let imageData = null;
  if (el.imgSrc) {
    try {
      let src = el.imgSrc;
      // puppeteer 在 file:// 下 el.src 通常已是 file:///... 绝对形式
      if (src.startsWith('file:///')) {
        // file:///C:/path/to/img.png  →  C:/path/to/img.png (Windows) 或 /path (Unix)
        src = src.replace(/^file:\/\/\//, '');
        // Windows: file:///C:/... 去掉 file:/// 后是 C:/...
        if (process.platform === 'win32' && /^[A-Za-z]:/.test(src) === false) {
          // Unix-like already
        }
        // URL decode（处理空格、中文）
        src = decodeURIComponent(src);
      } else if (/^https?:/.test(src)) {
        // 远程图片暂不处理（保持占位框）
        src = null;
      }
      if (src && fs.existsSync(src)) {
        const buf = fs.readFileSync(src);
        const ext = path.extname(src).slice(1).toLowerCase();
        const mime = ext === 'jpg' ? 'jpeg' : ext;
        imageData = `data:image/${mime};base64,${buf.toString('base64')}`;
      }
    } catch (err) {
      // 读取失败，回退占位框
      imageData = null;
    }
  }

  if (imageData) {
    try {
      slide.addImage({
        data: imageData,
        x,
        y,
        w: Math.max(w, 0.1),
        h: Math.max(h, 0.1),
      });
      return;
    } catch (err) {
      // addImage 失败（格式不支持等），回退占位框
    }
  }

  slide.addShape('rect', {
    x,
    y,
    w: Math.max(w, 0.5),
    h: Math.max(h, 0.5),
    fill: { type: 'solid', color: 'F0F0F0' },
    line: { color: '999999', width: 1 },
    rectRadius: 0,
  });

  // 在占位框中央添加"图片"文字
  slide.addText('[图片]', {
    x,
    y,
    w: Math.max(w, 0.5),
    h: Math.max(h, 0.5),
    fontSize: Math.min(14, Math.max(h * 72 * 0.3, 8)),
    fontFace: DEFAULT_CN_FONT,
    color: '999999',
    align: 'center',
    valign: 'middle',
  });
}

/**
 * 生成线条
 */
function generateLine(slide, el, scale) {
  const x = pxToInch(el.x, scale);
  const y = pxToInch(el.y + el.height / 2, scale);
  const w = pxToInch(el.width, scale);

  slide.addShape('line', {
    x,
    y,
    w: Math.max(w, MIN_SHAPE_SIZE_INCH),
    h: 0,
    line: { color: 'CCCCCC', width: 1 },
  });
}

/**
 * 生成容器（带背景色的 div）
 */
function generateContainer(slide, el, scale) {
  // 调试日志
  if (el.borderWidth > 0 || el.borderTopWidth > 0 || (el.width >= 17 && el.width <= 19 && el.height >= 17 && el.height <= 19)) {
    console.log(`[GENERATE-CONTAINER] tag=${el.tag}, text="${el.text}", borderWidth=${el.borderWidth}, borderColor=${el.borderColor}, borderTopWidth=${el.borderTopWidth}, borderRadius=${el.borderRadius}, x=${el.x}, y=${el.y}, w=${el.width}, h=${el.height}, className=${el.className}`);
  }

  // 自定义修改：跳过"全画布背景形状"——这类形状是 .slide 容器的产物，
  // 它的实际背景图已经通过 body url 分支 addImage 全幅铺了，再画矩形会盖住
  const PPT_W = 13.333, PPT_H = 7.5;
  const isFullSlideBg =
    el.tag === 'DIV' &&
    el.className === 'slide' &&
    Math.abs(pxToInch(el.width, scale) - PPT_W) < 0.05 &&
    Math.abs(pxToInch(el.height, scale) - PPT_H) < 0.05;
  if (isFullSlideBg) return;

  let bgColor = el.backgroundColor;
  const hasBackground = bgColor && bgColor !== 'transparent' && bgColor !== 'rgba(0, 0, 0, 0)';

  const x = pxToInch(el.x, scale);
  const y = pxToInch(el.y, scale);
  const w = pxToInch(el.width, scale);
  const h = pxToInch(el.height, scale);

  // 检查是否有任何边框（包括单独的边框属性）
  const hasAnyBorder = el.borderWidth > 0 || 
                       el.borderTopWidth > 0 || 
                       el.borderRightWidth > 0 || 
                       el.borderBottomWidth > 0 || 
                       el.borderLeftWidth > 0;
  
  // 调试日志：对于 deco-ring 和 deco-dots 打印详细信息
  if (el.tag === 'DECO-RING' || el.tag === 'DECO-DOTS' || el.borderWidth > 0) {
    console.log(`[GENERATE-CONTAINER-DEBUG] tag=${el.tag}, text="${el.text}", borderWidth=${el.borderWidth}, borderColor=${el.borderColor}, borderTopWidth=${el.borderTopWidth}, borderRightWidth=${el.borderRightWidth}, borderBottomWidth=${el.borderBottomWidth}, borderLeftWidth=${el.borderLeftWidth}, hasAnyBorder=${hasAnyBorder}, hasBackground=${hasBackground}`);
  }

  // 当有背景色或有边框时，生成形状
  if (hasBackground || hasAnyBorder) {
    // 处理背景色
    let fillColor = null;
    let fillTransparency = 0;
    let fillGradient = null; // 用于存储渐变背景
    if (hasBackground) {
      if (bgColor.includes('gradient')) {
        fillGradient = parseGradient(bgColor);
        if (!fillGradient) {
          fillColor = gradientToColor(bgColor);
        }
      } else {
        fillColor = colorToHex(bgColor);
      }
      const transparency = extractTransparency(el.backgroundColor);
      if (transparency !== null && transparency > 0) {
        fillTransparency = transparency;
      }
    }

    // 当有背景色时生成填充矩形（即使透明度为1也设置）
    // 当只有边框没有背景色时，使用 line 属性来画边框
    if (hasBackground && (fillColor || fillGradient)) {
      const shapeOptions = {
        x,
        y,
        w: Math.max(w, MIN_SHAPE_SIZE_INCH),
        h: Math.max(h, MIN_SHAPE_SIZE_INCH),
      };
      // CSS transform: rotate(...) → PptxGenJS 原生 rotate 字段（单位：度）
      if (el.rotation) shapeOptions.rotate = el.rotation;

      if (fillGradient) {
        // 优先走"原生渐变"路径（PptxGenJS 生成后由 gradient-postprocess 替换为 <a:gradFill>）
        const gradCtx = slide.__html2pptxGradCtx;
        if (gradCtx) {
          const nativeOptions = {
            x,
            y,
            w: Math.max(w, MIN_SHAPE_SIZE_INCH),
            h: Math.max(h, MIN_SHAPE_SIZE_INCH),
          };
          if (el.rotation) nativeOptions.rotate = el.rotation;
          // 处理边框（与原 shapeOptions 保持一致）
          if (el.borderWidth > 0 && el.borderColor) {
            const borderColor = colorToHex(el.borderColor);
            nativeOptions.line = { color: borderColor, width: el.borderWidth * 0.75 };
            const borderTransparency = extractTransparency(el.borderColor);
            if (borderTransparency !== null && borderTransparency > 0) {
              nativeOptions.line.transparency = borderTransparency;
            }
          }
          // 处理圆角
          if (el.borderRadius > 0) {
            nativeOptions.rectRadius = pxToInch(el.borderRadius, scale);
          }
          const nativeShapeType = el.borderRadius > 0 ? 'roundRect' : 'rect';
          addNativeGradientShape(gradCtx, nativeShapeType, nativeOptions, fillGradient);
          // 已经直接调用 addShape，跳过下方通用的 addShape 流程
          return;
        }
        // 兜底（理论上不会走到，slide.__html2pptxGradCtx 默认始终存在）：
        // 把渐变首色作为纯色填充，至少保留色彩信息
        shapeOptions.fill = { type: 'solid', color: gradientToColor(bgColor) };
        if (fillTransparency > 0) {
          shapeOptions.fill.transparency = fillTransparency;
        }
      } else {
        shapeOptions.fill = { type: 'solid', color: fillColor };
        if (fillTransparency > 0) {
          shapeOptions.fill.transparency = fillTransparency;
        }
      }

      // 处理统一的边框（如果有）
      if (el.borderWidth > 0 && el.borderColor) {
        const borderColor = colorToHex(el.borderColor);
        shapeOptions.line = { color: borderColor, width: el.borderWidth * 0.75 };
        // 处理边框透明度（如果边框颜色有 alpha）
        // extractTransparency 返回 0-100 范围的值，PptxGenJS 也使用 0-100 范围
        const borderTransparency = extractTransparency(el.borderColor);
        if (borderTransparency !== null && borderTransparency > 0) {
          shapeOptions.line.transparency = borderTransparency;
        }
      }

      // 处理圆角
      if (el.borderRadius > 0) {
        shapeOptions.rectRadius = pxToInch(el.borderRadius, scale);
      }

      // 有圆角用 roundRect，无圆角用 rect
      const shapeType = el.borderRadius > 0 ? 'roundRect' : 'rect';
      slide.addShape(shapeType, shapeOptions);
    } else if (hasAnyBorder) {
      // 只有边框没有背景色时，使用 line 属性来画边框
      // 生成四条边框线
      const borderColor = el.borderColor ? colorToHex(el.borderColor) : '000000';
      const borderWidth = el.borderWidth > 0 ? el.borderWidth : (el.borderTopWidth || el.borderRightWidth || el.borderBottomWidth || el.borderLeftWidth);
      const borderTransparency = extractTransparency(el.borderColor);
      
      console.log(`[GENERATE-CONTAINER-BORDER] borderColor=${borderColor}, borderWidth=${borderWidth}, borderTransparency=${borderTransparency}, hasBorder=${hasAnyBorder}, hasBackground=${hasBackground}`);
      
      if (el.borderRadius > 0) {
        // 有圆角时，使用带边框的矩形（设置背景透明）
        const shapeOptions = {
          x,
          y,
          w: Math.max(w, MIN_SHAPE_SIZE_INCH),
          h: Math.max(h, MIN_SHAPE_SIZE_INCH),
          fill: { type: 'solid', color: 'FFFFFF', transparency: 100 }, // 完全透明背景
          line: { color: borderColor, width: borderWidth * 0.75 },
          rectRadius: pxToInch(el.borderRadius, scale),
        };
        if (el.rotation) shapeOptions.rotate = el.rotation;
        if (borderTransparency !== null && borderTransparency > 0) {
          shapeOptions.line.transparency = borderTransparency;
        }
        slide.addShape('roundRect', shapeOptions);
      } else {
        // 无圆角时，生成四条单独的边框线
        // 上边框
        if (el.borderTopWidth > 0 && el.borderTopStyle !== 'none' && el.borderTopStyle !== 'hidden') {
          const lineOpts = { color: colorToHex(el.borderTopColor) || borderColor, width: el.borderTopWidth * 0.75 };
          if (borderTransparency !== null && borderTransparency > 0) {
            lineOpts.transparency = borderTransparency;
          }
          slide.addShape('line', {
            x: x,
            y: y,
            w: w,
            h: 0,
            line: lineOpts,
          });
        } else if (el.borderWidth > 0 && el.borderColor) {
          const lineOpts = { color: borderColor, width: borderWidth * 0.75 };
          if (borderTransparency !== null && borderTransparency > 0) {
            lineOpts.transparency = borderTransparency;
          }
          slide.addShape('line', {
            x: x,
            y: y,
            w: w,
            h: 0,
            line: lineOpts,
          });
        }
        // 下边框
        if (el.borderBottomWidth > 0 && el.borderBottomStyle !== 'none' && el.borderBottomStyle !== 'hidden') {
          const lineOpts = { color: colorToHex(el.borderBottomColor) || borderColor, width: el.borderBottomWidth * 0.75 };
          if (borderTransparency !== null && borderTransparency > 0) {
            lineOpts.transparency = borderTransparency;
          }
          slide.addShape('line', {
            x: x,
            y: y + h,
            w: w,
            h: 0,
            line: lineOpts,
          });
        } else if (el.borderWidth > 0 && el.borderColor) {
          const lineOpts = { color: borderColor, width: borderWidth * 0.75 };
          if (borderTransparency !== null && borderTransparency > 0) {
            lineOpts.transparency = borderTransparency;
          }
          slide.addShape('line', {
            x: x,
            y: y + h,
            w: w,
            h: 0,
            line: lineOpts,
          });
        }
        // 左边框
        if (el.borderLeftWidth > 0 && el.borderLeftStyle !== 'none' && el.borderLeftStyle !== 'hidden') {
          const lineOpts = { color: colorToHex(el.borderLeftColor) || borderColor, width: el.borderLeftWidth * 0.75 };
          if (borderTransparency !== null && borderTransparency > 0) {
            lineOpts.transparency = borderTransparency;
          }
          slide.addShape('line', {
            x: x,
            y: y,
            w: 0,
            h: h,
            line: lineOpts,
          });
        } else if (el.borderWidth > 0 && el.borderColor) {
          const lineOpts = { color: borderColor, width: borderWidth * 0.75 };
          if (borderTransparency !== null && borderTransparency > 0) {
            lineOpts.transparency = borderTransparency;
          }
          slide.addShape('line', {
            x: x,
            y: y,
            w: 0,
            h: h,
            line: lineOpts,
          });
        }
        // 右边框
        if (el.borderRightWidth > 0 && el.borderRightStyle !== 'none' && el.borderRightStyle !== 'hidden') {
          const lineOpts = { color: colorToHex(el.borderRightColor) || borderColor, width: el.borderRightWidth * 0.75 };
          if (borderTransparency !== null && borderTransparency > 0) {
            lineOpts.transparency = borderTransparency;
          }
          slide.addShape('line', {
            x: x + w,
            y: y,
            w: 0,
            h: h,
            line: lineOpts,
          });
        } else if (el.borderWidth > 0 && el.borderColor) {
          const lineOpts = { color: borderColor, width: borderWidth * 0.75 };
          if (borderTransparency !== null && borderTransparency > 0) {
            lineOpts.transparency = borderTransparency;
          }
          slide.addShape('line', {
            x: x + w,
            y: y,
            w: 0,
            h: h,
            line: lineOpts,
          });
        }
      }
    }
  }
}

/**
 * 生成圆形/椭圆
 */
function generateCircle(slide, el, scale) {
  // 调试日志
  console.log(`[GENERATE-CIRCLE] tag=${el.tag}, text="${el.text}", borderWidth=${el.borderWidth}, borderColor=${el.borderColor}, backgroundColor=${el.backgroundColor}, x=${el.x}, y=${el.y}, w=${el.width}, h=${el.height}`);
  
  const x = pxToInch(el.x, scale);
  const y = pxToInch(el.y, scale);
  const w = pxToInch(el.width, scale);
  const h = pxToInch(el.height, scale);

  const shapeOptions = {
    x,
    y,
    w: Math.max(w, MIN_SHAPE_SIZE_INCH),
    h: Math.max(h, MIN_SHAPE_SIZE_INCH),
  };
  if (el.rotation) shapeOptions.rotate = el.rotation;

  // 处理背景色
  let bgColor = el.backgroundColor;
  const hasBackground = bgColor && bgColor !== 'transparent' && bgColor !== 'rgba(0, 0, 0, 0)';

  // 渐变特殊路径：圆形/椭圆背景为线性渐变时，优先走"原生渐变"
  let circleGradient = null;
  if (hasBackground && bgColor.includes('gradient')) {
    circleGradient = parseGradient(bgColor);
  }
  const gradCtx = slide.__html2pptxGradCtx;
  if (circleGradient && gradCtx) {
    const nativeOptions = {
      x,
      y,
      w: Math.max(w, MIN_SHAPE_SIZE_INCH),
      h: Math.max(h, MIN_SHAPE_SIZE_INCH),
    };
    if (el.rotation) nativeOptions.rotate = el.rotation;
    // 处理边框（与下方原逻辑保持一致）
    const hasBorder = el.borderWidth > 0 || el.borderTopWidth > 0 ||
                      el.borderRightWidth > 0 || el.borderBottomWidth > 0 ||
                      el.borderLeftWidth > 0;
    if (hasBorder && el.borderColor) {
      const borderColor = colorToHex(el.borderColor);
      const borderWidth = el.borderWidth > 0 ? el.borderWidth :
                          (el.borderTopWidth || el.borderRightWidth || el.borderBottomWidth || el.borderLeftWidth);
      const borderTransparency = extractTransparency(el.borderColor);
      nativeOptions.line = { color: borderColor, width: borderWidth * 0.75 };
      if (borderTransparency !== null && borderTransparency > 0) {
        nativeOptions.line.transparency = borderTransparency;
      }
    }
    addNativeGradientShape(gradCtx, 'ellipse', nativeOptions, circleGradient);
    return;
  }

  if (hasBackground) {
    if (bgColor.includes('gradient')) {
      bgColor = gradientToColor(bgColor);
    } else {
      bgColor = colorToHex(bgColor);
    }
    if (bgColor) {
      shapeOptions.fill = { type: 'solid', color: bgColor };
      // 处理透明度（extractTransparency 返回 0-100 范围）
      const transparency = extractTransparency(el.backgroundColor);
      if (transparency !== null && transparency > 0) {
        shapeOptions.fill.transparency = transparency;
      }
    }
  }

  // 处理边框
  const hasBorder = el.borderWidth > 0 || el.borderTopWidth > 0 || 
                    el.borderRightWidth > 0 || el.borderBottomWidth > 0 || 
                    el.borderLeftWidth > 0;
  if (hasBorder && el.borderColor) {
    const borderColor = colorToHex(el.borderColor);
    const borderWidth = el.borderWidth > 0 ? el.borderWidth : 
                        (el.borderTopWidth || el.borderRightWidth || el.borderBottomWidth || el.borderLeftWidth);
    const borderTransparency = extractTransparency(el.borderColor);
    console.log(`[GENERATE-CIRCLE-BORDER] borderColor=${el.borderColor}, borderTransparency=${borderTransparency}`);
    shapeOptions.line = { color: borderColor, width: borderWidth * 0.75 };
    if (borderTransparency !== null && borderTransparency > 0) {
      shapeOptions.line.transparency = borderTransparency;
      console.log(`[GENERATE-CIRCLE-BORDER] 设置边框透明度=${borderTransparency}`);
    }
  }

  // 圆形必须设置填充，即使是透明的（否则 PptxGenJS 可能无法正确渲染边框）
  if (!shapeOptions.fill) {
    shapeOptions.fill = { type: 'solid', color: 'FFFFFF', transparency: 100 };
  }

  // 圆形用 ellipse，PptxGenJS 会根据宽高自动变成圆或椭圆
  slide.addShape('ellipse', shapeOptions);
}

/**
 * 生成 SVG 图片（Base64 嵌入）
 */
function generateSvgImage(slide, svg, scale) {
  const x = pxToInch(svg.x, scale);
  const y = pxToInch(svg.y, scale);
  const w = pxToInch(svg.width, scale);
  const h = pxToInch(svg.height, scale);

  try {
    slide.addImage({
      data: svg.dataUrl,
      x,
      y,
      w: Math.max(w, MIN_SHAPE_SIZE_INCH),
      h: Math.max(h, MIN_SHAPE_SIZE_INCH),
    });
  } catch (err) {
    // SVG 嵌入失败则生成占位框
    slide.addShape('rect', {
      x, y, w: Math.max(w, MIN_SHAPE_SIZE_INCH), h: Math.max(h, MIN_SHAPE_SIZE_INCH),
      fill: { type: 'solid', color: 'F0F0F0' },
      line: { color: 'CCCCCC', width: 1 },
    });
    slide.addText('[SVG]', { x, y, w, h, align: 'center', valign: 'middle', fontSize: 12, color: '999999' });
  }
}

/**
 * 生成伪元素（方案 A 实现）
 *
 * 设计思路：
 *   ① 形状层 —— 把伪元素"伪装"成一个普通元素对象 pseudoAsEl，
 *      直接调 generateContainer，从而免费获得：原生渐变（gradFill）、边框、
 *      圆角、透明度、形状选择（rect/roundRect）等完整能力。
 *   ② 文字层 —— 若 content 非空，则像原来一样用 addText 渲染。
 *      ✓ 等单字符图标当前继续走原硬编码偏移规则（方案 D 单独迭代时再泛化）。
 *
 * 为什么这样改：原 generatePseudoElement 重新实现了一遍简化版形状逻辑，
 * 不支持渐变、边框、透明度，导致 H5 中"渐变光带"、"带透明度的横线"等装饰
 * 在 PPT 里只剩首色纯色矩形，视觉差距明显。
 */
function generatePseudoElement(slide, pseudo, scale) {
  // 1. 形状层：构造 pseudoAsEl 并复用 generateContainer
  //    只在存在背景色 / 渐变 / 边框时才生成形状（与原行为一致）
  const hasBg = pseudo.backgroundColor
    && pseudo.backgroundColor !== 'transparent'
    && pseudo.backgroundColor !== 'rgba(0, 0, 0, 0)';
  const hasBorder = (pseudo.borderWidth || 0) > 0
    || (pseudo.borderTopWidth || 0) > 0
    || (pseudo.borderRightWidth || 0) > 0
    || (pseudo.borderBottomWidth || 0) > 0
    || (pseudo.borderLeftWidth || 0) > 0;

  if (hasBg || hasBorder) {
    const pseudoAsEl = {
      tag: 'PSEUDO_AS_CONTAINER',
      text: '',
      x: pseudo.x,
      y: pseudo.y,
      width: pseudo.width,
      height: pseudo.height,
      backgroundColor: pseudo.backgroundColor,
      borderRadius: pseudo.borderRadius || 0,
      borderWidth: pseudo.borderWidth || 0,
      borderTopWidth: pseudo.borderTopWidth || 0,
      borderRightWidth: pseudo.borderRightWidth || 0,
      borderBottomWidth: pseudo.borderBottomWidth || 0,
      borderLeftWidth: pseudo.borderLeftWidth || 0,
      borderColor: pseudo.borderColor || '',
      // generateContainer 通过 extractTransparency(el.backgroundColor) 读取背景透明度，
      // 这里我们额外把伪元素自身的 opacity 折算进 backgroundColor 的 transparency 中：
      // 若 opacity < 1 且当前 backgroundColor 是 rgb/hex/gradient 字符串，直接传给后续流程；
      // generateContainer 内部对 gradient 走 parseGradient，色标透明度由 parseGradient 保留。
    };
    // 处理 opacity（< 1 时叠加到 fillTransparency）：
    // 简化策略 —— 若是纯色背景，且 opacity < 1，我们把颜色字符串转成 rgba(...) 形式，
    // 让下游 extractTransparency 能自动读出 (1-opacity)*100 的透明度。
    if (pseudo.opacity !== undefined && pseudo.opacity < 1 && hasBg
        && !pseudo.backgroundColor.includes('gradient')) {
      pseudoAsEl.backgroundColor = wrapColorWithOpacity(pseudo.backgroundColor, pseudo.opacity);
    }
    // 对渐变背景的 opacity，需通过后处理时把每个 stop 的 transparency 叠加，
    // 这里挂一个标记，generateContainer → addNativeGradientShape 链路目前不读取，
    // 后续若发现渐变光带颜色偏深，再扩展 gradient-postprocess 支持整体 opacity。
    if (pseudo.opacity !== undefined && pseudo.opacity < 1 && hasBg
        && pseudo.backgroundColor.includes('gradient')) {
      pseudoAsEl.__pseudoOpacity = pseudo.opacity;
    }

    generateContainer(slide, pseudoAsEl, scale);
  }

  // 2. 文字层：与原实现保持一致，仅修正字重/对齐取自伪元素自身
  if (pseudo.text && pseudo.text.trim()) {
    const x = pxToInch(pseudo.x, scale);
    const y = pxToInch(pseudo.y, scale);
    const w = pxToInch(pseudo.width, scale);
    const h = pxToInch(pseudo.height, scale);
    const fontSize = parseFontSize(pseudo.fontSize);
    const font = resolveFont(pseudo.text, pseudo.fontFamily);
    const color = colorToHex(pseudo.color) || '000000';

    let textX = x;
    let textY = y;
    let textW = Math.max(w, MIN_VISIBLE_TEXT_SIZE_INCH);
    let textH = Math.max(h, MIN_VISIBLE_TEXT_SIZE_INCH);

    // ✓ 图标专属偏移（保持向后兼容；后续方案 D 会泛化为单字符图标通用规则）
    if (pseudo.text === '✓') {
      textX = pxToInch(pseudo.x - 5.25, scale);
      textY = pxToInch(pseudo.y - 9, scale);
      textW = pxToInch(18, scale);
      textH = pxToInch(18, scale);
    }

    const textOptions = {
      x: textX,
      y: textY,
      w: textW,
      h: textH,
      fontSize,
      fontFace: font,
      color,
      align: 'center',
      valign: 'middle',
    };
    // 应用字重 / 字形 / 下划线（伪元素现在能正确表达 bold/italic 等）
    if (pseudo.fontWeight && (pseudo.fontWeight === 'bold' || parseInt(pseudo.fontWeight, 10) >= 600)) {
      textOptions.bold = true;
    }
    if (pseudo.fontStyle === 'italic') {
      textOptions.italic = true;
    }
    if (pseudo.textDecoration && pseudo.textDecoration.includes('underline')) {
      textOptions.underline = { style: 'sng' };
    }
    slide.addText(pseudo.text, textOptions);
  }
}

/**
 * 把任意颜色字符串 + opacity(0-1) 包装成等价的 rgba(...) 字符串
 * 仅用于"伪元素 opacity < 1 + 纯色背景"场景，让 extractTransparency 能正确读出透明度
 */
function wrapColorWithOpacity(colorStr, opacity) {
  if (!colorStr || typeof colorStr !== 'string') return colorStr;
  const alpha = Math.max(0, Math.min(1, opacity));
  // 已是 rgba(...) → 乘以 alpha
  const rgbaMatch = colorStr.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)$/i);
  if (rgbaMatch) {
    const r = rgbaMatch[1];
    const g = rgbaMatch[2];
    const b = rgbaMatch[3];
    const oldAlpha = rgbaMatch[4] !== undefined ? parseFloat(rgbaMatch[4]) : 1;
    return `rgba(${r}, ${g}, ${b}, ${(oldAlpha * alpha).toFixed(3)})`;
  }
  // #hex → 转 rgba
  const hexMatch = colorStr.match(/^#?([0-9a-f]{6})$/i);
  if (hexMatch) {
    const r = parseInt(hexMatch[1].slice(0, 2), 16);
    const g = parseInt(hexMatch[1].slice(2, 4), 16);
    const b = parseInt(hexMatch[1].slice(4, 6), 16);
    return `rgba(${r}, ${g}, ${b}, ${alpha.toFixed(3)})`;
  }
  // 其他不识别格式 → 原样返回（不致命，最多丢失 opacity）
  return colorStr;
}

/**
 * 生成表格
 */
function generateTable(slide, table, scale) {
  if (!table.rows || table.rows.length === 0) return;

  try {
    const x = pxToInch(table.x, scale);
    const y = pxToInch(table.y, scale);
    const w = pxToInch(table.width, scale);

    // 计算列数
    const colCount = Math.max(...table.rows.map(r => r.length));
    const colWidth = w / colCount;

    // 构建表格数据
    const tableRows = table.rows.map((row, rowIdx) => {
      return row.map((cell, colIdx) => {
        const cellColor = colorToHex(cell.color);
        const cellBg = cell.backgroundColor && cell.backgroundColor !== 'rgba(0, 0, 0, 0)' && cell.backgroundColor !== 'transparent'
          ? colorToHex(cell.backgroundColor)
          : null;
        const cellFont = resolveFont(cell.text || '', cell.fontFamily);
        const cellFontSize = parseFontSize(cell.fontSize);
        const cellBold = cell.fontWeight === 'bold' || parseInt(cell.fontWeight) >= 600;
        const cellAlign = parseAlignment(cell.textAlign);

        const cellData = {
          text: cell.text || '',
          options: {
            fontSize: cellFontSize,
            fontFace: cellFont,
            color: cellColor || '000000',
            bold: cellBold,
            align: cellAlign,
            valign: 'middle',
            margin: [3, 5, 3, 5],
          },
        };

        // 只在有背景色时添加 fill（避免 undefined 属性干扰）
        if (cellBg) {
          cellData.options.fill = { type: 'solid', color: cellBg };
        }

        return cellData;
      });
    });

    slide.addTable(tableRows, {
      x,
      y,
      w: Math.max(w, 0.5),
      colW: Array(colCount).fill(colWidth),
      rowH: table.rowHeights ? table.rowHeights.map(h => h / 96) : undefined,
      border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
      autoPage: false,
    });
  } catch (err) {
    console.error('[generateTable] 生成表格时出错:', err.message);
  }
}

/**
 * 生成完整的 PPT 文件
 * @param {Array<object>} slideDataList - 多页提取数据数组
 * @param {string} outputPath - 输出文件路径
 * @param {object} options - 选项
 */
async function generatePptx(slideDataList, outputPath, options = {}) {
  const { scale = DEFAULT_SCALE, verbose = false } = options;

  if (verbose) console.log(`[generate] 开始生成 PPT，共 ${slideDataList.length} 页`);

  // 创建 PPT 实例
  const pptx = new PptxGenJS();

  // 设置 PPT 属性
  pptx.layout = 'LAYOUT_WIDE'; // 16:9
  pptx.author = 'html2pptx';
  pptx.title = 'HTML to PPT Conversion';

  // 渐变任务收集器（PptxGenJS 3.12.0 不支持原生 gradFill API，
  // 这里收集所有需要的渐变信息，写文件后由 gradient-postprocess 改写 OOXML）
  const collector = createGradientCollector();

  // 逐页生成
  slideDataList.forEach((slideData, index) => {
    if (verbose) console.log(`[generate] 生成第 ${index + 1} 页...`);
    generateSlide(pptx, slideData, { scale, verbose, collector, slideIndex: index });
  });

  // 写入文件
  await pptx.writeFile({ fileName: outputPath });

  if (verbose) console.log(`[generate] PPT 已保存: ${outputPath}`);

  // 后处理：将占位形状的 solidFill 替换为 OOXML 原生 gradFill
  const tasks = collector.getTasks();
  if (tasks.length > 0) {
    if (verbose) console.log(`[generate] 开始渐变后处理，共 ${tasks.length} 个任务`);
    try {
      const { applied, skipped } = await applyGradients(outputPath, tasks, { verbose });
      if (verbose) console.log(`[generate] 渐变后处理完成：成功 ${applied}，跳过 ${skipped}`);
    } catch (err) {
      console.warn(`[generate] 渐变后处理失败（保留占位纯色）：${err.message}`);
    }
  }
}

module.exports = { generateSlide, generatePptx };
