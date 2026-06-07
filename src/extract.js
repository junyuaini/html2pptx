/**
 * html2pptx - Puppeteer 元素提取模块
 * 使用 Puppeteer 渲染 HTML 并提取所有可见元素的位置和样式
 */

const puppeteer = require('puppeteer');
const path = require('path');
const {
  colorToHex,
  extractTransparency,
  gradientToColor,
  parseFontSize,
  parseAlignment,
  normalizeFont,
  resolveFont,
  isTextElement,
  isContainerElement,
} = require('./utils');

/**
 * 从单个 HTML 文件提取所有可见元素
 * @param {string} htmlPath - HTML 文件的绝对路径
 * @param {object} options - 选项
 * @param {boolean} verbose - 是否输出详细日志
 * @returns {Promise<object>} 提取结果 { elements, tables, pageWidth, pageHeight }
 */
async function extractFromHtml(htmlPath, options = {}, verbose = false) {
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu'],
  });

  try {
    const page = await browser.newPage();

    // 监听浏览器控制台日志
    page.on('console', msg => {
      console.log(`[BROWSER] ${msg.text()}`);
    });

    // 设置视口为 H5 标准尺寸
    await page.setViewport({ width: 1280, height: 720 });

    // 打开 HTML 文件
    const fileUrl = path.isAbsolute(htmlPath)
      ? 'file://' + htmlPath
      : 'file://' + path.resolve(process.cwd(), htmlPath);

    if (verbose) console.log(`[extract] 打开文件: ${fileUrl}`);
    await page.goto(fileUrl, { waitUntil: 'networkidle0', timeout: 30000 });

    // 提取所有可见元素和页面信息
    const extracted = await page.evaluate(() => {
      /**
       * 公共变换解析（运行在浏览器上下文）：
       * 从 computed style 的 transform 字段反解 rotation / translate / scale。
       * computed 永远以 matrix(a,b,c,d,tx,ty) 形式返回（CSS3 规范行为），
       * % 单位的 translateX/Y 在 matrix 里已被换算成像素，无需再次处理。
       *
       * 数学约定（CSS3 标准）：
       *   matrix = [ a c tx ]   2D 仿射：(a,b)=旋转+缩放 X 分量
       *            [ b d ty ]              (c,d)=旋转+缩放 Y 分量
       *            [ 0 0 1  ]              (tx,ty)=平移
       *   rotation = atan2(b, a)（弧度）→ 转度
       *   scaleX = sqrt(a*a + b*b)
       *   scaleY = sqrt(c*c + d*d)
       */
      function parseTransform(style) {
        const result = { rotation: 0, translateX: 0, translateY: 0, scaleX: 1, scaleY: 1 };
        if (!style || !style.transform || style.transform === 'none') return result;
        const m2d = style.transform.match(/matrix\(\s*([-\d.]+),\s*([-\d.]+),\s*([-\d.]+),\s*([-\d.]+),\s*([-\d.]+),\s*([-\d.]+)\s*\)/);
        if (m2d) {
          const a = parseFloat(m2d[1]);
          const b = parseFloat(m2d[2]);
          const c = parseFloat(m2d[3]);
          const d = parseFloat(m2d[4]);
          result.translateX = parseFloat(m2d[5]) || 0;
          result.translateY = parseFloat(m2d[6]) || 0;
          result.rotation = Math.round(Math.atan2(b, a) * 180 / Math.PI);
          result.scaleX = Math.sqrt(a * a + b * b);
          result.scaleY = Math.sqrt(c * c + d * d);
          return result;
        }
        const m3d = style.transform.match(/matrix3d\(([-\d.,\s]+)\)/);
        if (m3d) {
          const v = m3d[1].split(',').map(s => parseFloat(s.trim()));
          if (v.length >= 16) {
            // matrix3d 列主序：v[0..3]=col1, v[4..7]=col2, ... v[12..14]=平移
            result.translateX = v[12] || 0;
            result.translateY = v[13] || 0;
            // 简化处理：仅取 2D 投影的旋转/缩放
            result.rotation = Math.round(Math.atan2(v[1], v[0]) * 180 / Math.PI);
            result.scaleX = Math.sqrt(v[0] * v[0] + v[1] * v[1]);
            result.scaleY = Math.sqrt(v[4] * v[4] + v[5] * v[5]);
          }
        }
        return result;
      }

      const body = document.body;
      const firstChild = body.firstElementChild;
      const pageInfo = {
        width: firstChild ? firstChild.offsetWidth : body.offsetWidth,
        height: firstChild ? firstChild.offsetHeight : body.offsetHeight,
      };
      // 提取 body 元素的背景色
      const bodyStyle = window.getComputedStyle(body);
      let bodyBackgroundColor = bodyStyle.backgroundColor;
      if (bodyBackgroundColor === 'rgba(0, 0, 0, 0)' || bodyBackgroundColor === 'transparent') {
        const bgImage = bodyStyle.backgroundImage;
        if (bgImage && bgImage !== 'none') {
          bodyBackgroundColor = bgImage;
        } else {
          bodyBackgroundColor = null;
        }
      }
      const results = [];
      const allElements = document.querySelectorAll('body *');

      allElements.forEach(el => {
        const rect = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        const tag = el.tagName;
        const elClass = el.className || '';
        
        // 跳过隐藏元素
        if (style.display === 'none' || style.visibility === 'hidden') return;
        if (rect.width < 1 && rect.height < 1) return;

        // 提取文本内容（所有元素都尝试提取）
        let text = '';
        const directTextTags = ['H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'P', 'LI', 'SPAN', 'A', 'B', 'STRONG', 'I', 'EM', 'U', 'TD', 'TH', 'CAPTION'];
        if (directTextTags.includes(tag)) {
          text = el.innerText || el.textContent || '';
        } else {
          // 对于 DIV 等容器，只提取直接文本子节点（避免重复提取子元素文字）
          let directText = '';
          el.childNodes.forEach(node => {
            if (node.nodeType === 3) { // 文本节点
              directText += node.textContent;
            }
          });
          text = directText.trim();
        }
        // 清理文本：去除首尾空白，并清理内部的换行和多余空格
        text = text.trim().replace(/\s+/g, ' ');

        // 提取富文本 runs（用于文本类元素）
        let runs = null;
        if (['H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'P', 'LI', 'A', 'TD', 'TH'].includes(tag)) {
          runs = extractRuns(el);
        } else if (tag === 'DIV' && text && text.length > 0) {
          // DIV 有文字内容时也提取 runs（如 .adv-text 等多行文本 DIV）
          runs = extractRuns(el);
        }

        // 提取图片信息
        let imgSrc = null;
        if (tag === 'IMG') {
          imgSrc = el.src || el.getAttribute('src');
        }

        // 提取背景色
        let bgColor = style.backgroundColor;
        // 如果背景色是透明的，但有背景图片或渐变
        if (bgColor === 'rgba(0, 0, 0, 0)' || bgColor === 'transparent') {
          const bgImage = style.backgroundImage;
          if (bgImage && bgImage !== 'none') {
            bgColor = bgImage; // 保留原始值，后续处理渐变
          } else {
            bgColor = null;
          }
        }

        // 检测是否是圆形（border-radius: 50% 或宽高相等且圆角足够大）
        const borderRadiusRaw = style.borderRadius;
        const isCircle = borderRadiusRaw === '50%' ||
          (rect.width === rect.height && parseFloat(borderRadiusRaw) >= rect.width / 2);
        // 调试日志：对于装饰类元素打印圆形判断信息
        if (elClass.toString().includes('deco-ring')) {
          console.log(`[EXTRACT-CIRCLE] class=deco-ring, borderRadiusRaw=${borderRadiusRaw}, rect.width=${rect.width}, rect.height=${rect.height}, parseFloat(borderRadiusRaw)=${parseFloat(borderRadiusRaw)}, rect.width/2=${rect.width/2}, isCircle=${isCircle}`);
        }

        // CSS transform 解析 + 位置/尺寸修正（D1 方案）
        // 浏览器 getBoundingClientRect 返回的是变换后的外接矩形（AABB），
        // 对旋转元素需用 CSS 声明的原始 width/height，并反推未变换前的左上角。
        const transform = parseTransform(style);
        const cssW = parseFloat(style.width) || 0;
        const cssH = parseFloat(style.height) || 0;
        let finalX = rect.left;
        let finalY = rect.top;
        let finalWidth = rect.width;
        let finalHeight = rect.height;
        // 只有非零旋转时修正（避免对 translate-only / scale-only 等不改变 AABB 的场景做无意义修正）
        if (transform.rotation !== 0 && cssW > 0 && cssH > 0) {
          // AABB 中心与未旋转元素中心相同
          const centerX = rect.left + rect.width / 2;
          const centerY = rect.top + rect.height / 2;
          finalWidth = cssW;
          finalHeight = cssH;
          finalX = centerX - cssW / 2;
          finalY = centerY - cssH / 2;
        }

        results.push({
          tag,
          text,
          className: el.className || '',
          runs,
          imgSrc,
          x: finalX,
          y: finalY,
          width: finalWidth,
          height: finalHeight,
          color: style.color,
          fontSize: style.fontSize,
          fontFamily: style.fontFamily,
          fontWeight: style.fontWeight,
          fontStyle: style.fontStyle,
          textDecoration: style.textDecoration,
          textAlign: style.textAlign,
          lineHeight: style.lineHeight,
          backgroundColor: bgColor,
          borderColor: style.borderColor,
          borderWidth: parseFloat(style.borderWidth) || 0,
          borderStyle: style.borderStyle,
          borderRadius: parseFloat(style.borderRadius) || 0,
          // 单独的边框属性（用于坐标轴等）
          borderTopWidth: parseFloat(style.borderTopWidth) || 0,
          borderRightWidth: parseFloat(style.borderRightWidth) || 0,
          borderBottomWidth: parseFloat(style.borderBottomWidth) || 0,
          borderLeftWidth: parseFloat(style.borderLeftWidth) || 0,
          borderTopColor: style.borderTopColor,
          borderRightColor: style.borderRightColor,
          borderBottomColor: style.borderBottomColor,
          borderLeftColor: style.borderLeftColor,
          borderTopStyle: style.borderTopStyle,
          borderRightStyle: style.borderRightStyle,
          borderBottomStyle: style.borderBottomStyle,
          borderLeftStyle: style.borderLeftStyle,
          isCircle,
          opacity: parseFloat(style.opacity),
          // CSS transform 信息（仅 rotation 当前会被 generate 阶段使用）
          rotation: transform.rotation,
          padding: {
            top: parseFloat(style.paddingTop) || 0,
            right: parseFloat(style.paddingRight) || 0,
            bottom: parseFloat(style.paddingBottom) || 0,
            left: parseFloat(style.paddingLeft) || 0,
          },
        });
      });

      // 调试：打印所有找到的元素
      console.log('--- ALL ELEMENTS ---');
      results.forEach(el => {
        if (el.className && el.className.includes('check')) {
          console.log(`[ELEMENT] ${el.className} at x=${el.x}, y=${el.y}, w=${el.width}, h=${el.height}, bg=${el.backgroundColor}, bw=${el.borderWidth}`);
        }
      });
      console.log('--- END ELEMENTS ---');

      // 调试日志：检查边框相关元素（在 results 数组中检查）
      const borderRelatedClasses = ['top-module', 'business-card', 'bottom-card', 'base-section', 'ai-section', 'ai-item'];
      results.forEach(r => {
        // 检查有边框的元素
        if (r.borderWidth > 0 || r.borderTopWidth > 0 || r.borderRightWidth > 0 || r.borderBottomWidth > 0 || r.borderLeftWidth > 0) {
          console.log(`[EXTRACT-RESULT] tag=${r.tag}, text="${r.text}", borderWidth=${r.borderWidth}, borderColor=${r.borderColor}, borderTopWidth=${r.borderTopWidth}, borderRightWidth=${r.borderRightWidth}, borderBottomWidth=${r.borderBottomWidth}, borderLeftWidth=${r.borderLeftWidth}`);
        }
      });

      // 注意：不再进行容器高度扩展
      // Puppeteer提取的是浏览器渲染后的实际位置和尺寸，已经是准确的
      // 任何后续修改都会破坏原始布局，导致PPT与H5不一致

      // 收集调试信息：关键元素位置
      const debugElements = results.filter(el => 
        el.backgroundColor && el.backgroundColor.includes('f5f7fa')
      );
      const debugInfo = [];
      if (debugElements.length > 0) {
        debugElements.forEach((el, idx) => {
          const hasText = el.text && el.text.trim();
          const textPreview = hasText ? el.text.substring(0, 20) + '...' : '无文本';
          debugInfo.push(`元素 ${idx+1}: tag=${el.tag}, x=${el.x}, y=${el.y}, w=${el.width}, h=${el.height}, 文本="${textPreview}"`);
        });
      }

      // 提取 SVG 元素（转为图片数据）
      const svgs = [];
      document.querySelectorAll('svg').forEach(svg => {
        const rect = svg.getBoundingClientRect();
        try {
          const svgData = new XMLSerializer().serializeToString(svg);
          const svgBase64 = btoa(unescape(encodeURIComponent(svgData)));
          svgs.push({
            x: rect.left,
            y: rect.top,
            width: rect.width,
            height: rect.height,
            dataUrl: 'data:image/svg+xml;base64,' + svgBase64,
          });
        } catch (e) {
          // SVG 序列化失败则跳过
        }
      });

      // 提取伪元素（::before / ::after）
      const pseudoElements = [];

      /**
       * 在 document.styleSheets 中查找首个匹配 (el + pseudoType) 的 CSSRule，
       * 返回其 style 对象；找不到或访问受限时返回 null。
       *
       * 设计目的：把原本 4 层嵌套 (try → for → try → for) 抽成独立函数，
       * 让主循环只关心"拿到 rule.style 后怎么用"，提升可读性与可维护性。
       *
       * 注意：本函数运行在浏览器上下文（page.evaluate 内），不能依赖 Node API。
       */
      function findPseudoRuleFromStylesheets(el, pseudoType) {
        let result = null;
        try {
          for (let i = 0; i < document.styleSheets.length; i++) {
            const sheet = document.styleSheets[i];
            let rules = null;
            try {
              rules = sheet.cssRules || sheet.rules;
            } catch (e) {
              // 跨域样式表访问会抛 SecurityError —— 静默跳过
              continue;
            }
            if (!rules) continue;
            for (let j = 0; j < rules.length; j++) {
              const rule = rules[j];
              if (!rule.selectorText) continue;
              if (!rule.selectorText.includes(pseudoType)) continue;
              const baseSelector = rule.selectorText.replace(pseudoType, '').trim();
              try {
                if (el.matches(baseSelector)) {
                  result = rule.style;
                  break;
                }
              } catch (e) {
                // 无效选择器（少见，例如 ::deep 等非标准伪类）—— 跳过
              }
            }
            if (result) break;
          }
        } catch (e) {
          // styleSheets 本身不可访问的极端环境兜底
        }
        return result;
      }

      // 检查每个元素的伪元素
      const allPseudoCheckElements = document.querySelectorAll('*');
      
      allPseudoCheckElements.forEach(el => {
        const parentRect = el.getBoundingClientRect();
        
        // 检查 ::before 和 ::after
        ['::before', '::after'].forEach(pseudoType => {
          const style = window.getComputedStyle(el, pseudoType);
          
          const content = style.content;
          const hasBg = style.backgroundColor && style.backgroundColor !== 'rgba(0, 0, 0, 0)' && style.backgroundColor !== 'transparent';
          
          // 获取尺寸
          let w = parseFloat(style.width) || 0;
          let h = parseFloat(style.height) || 0;
          
          // 尝试从CSS规则中获取原始值（特别是百分比，针对伪元素）
          let foundRuleStyle = null;
          if (w === 0 || h === 0) {
            foundRuleStyle = findPseudoRuleFromStylesheets(el, pseudoType);
            if (foundRuleStyle) {
              if (w === 0 && foundRuleStyle.width) {
                w = foundRuleStyle.width.includes('%')
                  ? parentRect.width * parseFloat(foundRuleStyle.width) / 100
                  : parseFloat(foundRuleStyle.width) || 0;
              }
              if (h === 0 && foundRuleStyle.height) {
                h = foundRuleStyle.height.includes('%')
                  ? parentRect.height * parseFloat(foundRuleStyle.height) / 100
                  : parseFloat(foundRuleStyle.height) || 0;
              }
            }
          }
          
          // 处理绝对定位占满父元素的特殊装饰条情况
          if ((hasBg || (content && content !== 'none' && content !== '')) && (w === 0 || h === 0)) {
            const position = style.position;
            if (position === 'absolute' || position === 'fixed') {
              const top = style.top;
              const bottom = style.bottom;
              const left = style.left;
              const right = style.right;
              
              if (h === 0 && ((top === '0px' && bottom === '0px') || (top === '0' && bottom === '0'))) {
                h = parentRect.height;
              }
              if (w === 0 && ((left === '0px' && right === '0px') || (left === '0' && right === '0'))) {
                w = parentRect.width;
              }
            }
          }
          
          if ((hasBg || (w > 0 && h > 0) || (content && content !== 'none' && content !== '')) && (w > 0 || h > 0)) {
            // 计算伪元素位置
            let px = parentRect.left;
            let py = parentRect.top;
            
            const position = style.position;
            
            // 方案1: 对于绝对定位的伪元素，使用left/top/bottom/right计算
            if (position === 'absolute' || position === 'relative' || position === 'fixed') {
              px = parentRect.left;
              py = parentRect.top;

              // 标准 left / top
              if (style.left && style.left !== 'auto') {
                px += style.left.includes('%') ? parentRect.width * parseFloat(style.left)/100 : parseFloat(style.left) || 0;
              }
              if (style.top && style.top !== 'auto') {
                py += style.top.includes('%') ? parentRect.height * parseFloat(style.top)/100 : parseFloat(style.top) || 0;
              }

              // bottom / right 补位（当对应的 top/left 未设置时使用）
              if ((!style.top || style.top === 'auto') && style.bottom && style.bottom !== 'auto') {
                const b = style.bottom.includes('%')
                  ? parentRect.height * parseFloat(style.bottom) / 100
                  : parseFloat(style.bottom) || 0;
                py = parentRect.top + parentRect.height - b - h;
              }
              if ((!style.left || style.left === 'auto') && style.right && style.right !== 'auto') {
                const r = style.right.includes('%')
                  ? parentRect.width * parseFloat(style.right) / 100
                  : parseFloat(style.right) || 0;
                px = parentRect.left + parentRect.width - r - w;
              }

              // transform 解析：computed style 返回 matrix(a,b,c,d,tx,ty)
              // translateX(-50%) 已被浏览器解算为具体像素值写在 tx 上
              if (style.transform && style.transform !== 'none') {
                const m2d = style.transform.match(/matrix\(\s*[-\d.]+,\s*[-\d.]+,\s*[-\d.]+,\s*[-\d.]+,\s*([-\d.]+),\s*([-\d.]+)\s*\)/);
                if (m2d) {
                  px += parseFloat(m2d[1]) || 0;
                  py += parseFloat(m2d[2]) || 0;
                } else {
                  const m3d = style.transform.match(/matrix3d\([-\d.,\s]+,\s*([-\d.]+),\s*([-\d.]+),\s*[-\d.]+,\s*[-\d.]+\s*\)/);
                  if (m3d) {
                    px += parseFloat(m3d[1]) || 0;
                    py += parseFloat(m3d[2]) || 0;
                  }
                }
              }

              // 用CSS规则补充（仅当 computed style 未给出有效 left/top 时）：
              // 历史背景 —— foundRuleStyle 仅在 w===0 || h===0 时从样式表回填，
              // 早期作为"覆盖式"二次计算存在；但这与 transform / bottom / right
              // 的解析结果存在冲突（rule 覆盖会重置上面累计的偏移）。
              // 现在改为"补位式"：仅在浏览器未提供有效 left/top 时才使用 rule。
              if (foundRuleStyle) {
                const computedLeftMissing = !style.left || style.left === 'auto';
                const computedTopMissing = !style.top || style.top === 'auto';
                if (computedLeftMissing && foundRuleStyle.left && foundRuleStyle.left !== 'auto') {
                  px = parentRect.left + (foundRuleStyle.left.includes('%')
                    ? parentRect.width * parseFloat(foundRuleStyle.left) / 100
                    : parseFloat(foundRuleStyle.left) || 0);
                }
                if (computedTopMissing && foundRuleStyle.top && foundRuleStyle.top !== 'auto') {
                  py = parentRect.top + (foundRuleStyle.top.includes('%')
                    ? parentRect.height * parseFloat(foundRuleStyle.top) / 100
                    : parseFloat(foundRuleStyle.top) || 0);
                }
              }
            } else {
              // 方案2: 对于static定位的::before，位置 = 父元素left + 父元素padding-left + ::before margin-left
              // 对于::after，根据display处理
              if (pseudoType === '::before') {
                // ::before在文档流中的位置：父元素左边 + 父元素padding-left + ::before margin-left
                const parentStyle = window.getComputedStyle(el);
                const parentPaddingLeft = parseFloat(parentStyle.paddingLeft) || 0;
                const beforeMarginLeft = parseFloat(style.marginLeft) || 0;
                px = parentRect.left + parentPaddingLeft + beforeMarginLeft;
                
                // 对于flex容器，::before会被align-items: center垂直居中
                // 所以py需要调整
                const alignItems = parentStyle.alignItems;
                if (alignItems === 'center') {
                  // ::before垂直居中
                  py = parentRect.top + (parentRect.height - h) / 2;
                } else {
                  py = parentRect.top;
                }
              } else if (pseudoType === '::after') {
                // ::after的位置取决于display
                const display = style.display;
                if (display === 'block') {
                  // block的::after在父元素下方
                  try {
                    const tempSpan = document.createElement('span');
                    tempSpan.style.cssText = 'position:relative;display:inline-block;width:0;height:0;visibility:hidden;';
                    el.appendChild(tempSpan);
                    const tempRect = tempSpan.getBoundingClientRect();
                    px = tempRect.left;
                    py = tempRect.top;
                    el.removeChild(tempSpan);
                  } catch (e) {
                    px = parentRect.left;
                    py = parentRect.top + parentRect.height - h;
                  }
                } else {
                  // inline的::after在内容之后
                  try {
                    const tempSpan = document.createElement('span');
                    tempSpan.style.cssText = 'position:relative;display:inline;width:0;height:0;visibility:hidden;';
                    el.appendChild(tempSpan);
                    const tempRect = tempSpan.getBoundingClientRect();
                    px = tempRect.left;
                    py = tempRect.top;
                    el.removeChild(tempSpan);
                  } catch (e) {
                    px = parentRect.left + parentRect.width - w;
                    py = parentRect.top;
                  }
                }
              }
            }
            
            // 准备最终值
            const bgColor = style.backgroundColor || 'transparent';
            const borderRadius = parseFloat(style.borderRadius) || 0;
            const isDecorLine = pseudoType === '::before' && w > 0 && w < 15 && h > 0 && h >= parentRect.height * 0.5;
            
            // 提取伪元素的margin值（用于文本偏移计算）
            const marginLeft = parseFloat(style.marginLeft) || 0;
            const marginRight = parseFloat(style.marginRight) || 0;

            // 方案A：补齐伪元素属性，让 generate 阶段能复用 generateContainer 的完整能力
            // 1) 边框（伪元素也可能有 border）
            const borderWidth = parseFloat(style.borderWidth) || 0;
            const borderTopWidth = parseFloat(style.borderTopWidth) || 0;
            const borderRightWidth = parseFloat(style.borderRightWidth) || 0;
            const borderBottomWidth = parseFloat(style.borderBottomWidth) || 0;
            const borderLeftWidth = parseFloat(style.borderLeftWidth) || 0;
            const borderColor = style.borderColor || style.borderTopColor || '';
            // 2) 透明度（伪元素经常配合 opacity 做光晕效果）
            const opacity = parseFloat(style.opacity);
            // 3) 文本相关属性（伪元素可能注入文字内容并配合特定字重/对齐）
            const fontWeight = style.fontWeight || 'normal';
            const fontStyle = style.fontStyle || 'normal';
            const textAlign = style.textAlign || 'center';
            const textDecoration = style.textDecorationLine || style.textDecoration || 'none';
            // 4) 背景渐变图片单独标记（CSS 中渐变其实是 background-image，
            //    但 backgroundColor 在浏览器里有时也能拿到，这里二者兼取，generate 阶段优先用含 gradient 的那个）
            const bgImage = style.backgroundImage || 'none';
            const finalBackgroundColor = (bgImage && bgImage !== 'none' && bgImage.includes('gradient'))
              ? bgImage
              : bgColor;
            
            pseudoElements.push({
              tag: 'PSEUDO',
              text: content ? content.replace(/["']/g, '') : '',
              x: px,
              y: py,
              width: w,
              height: h,
              backgroundColor: finalBackgroundColor,
              color: style.color || 'inherit',
              fontSize: style.fontSize,
              fontFamily: style.fontFamily,
              fontWeight: fontWeight,
              fontStyle: fontStyle,
              textAlign: textAlign,
              textDecoration: textDecoration,
              borderRadius: borderRadius,
              borderWidth: borderWidth,
              borderTopWidth: borderTopWidth,
              borderRightWidth: borderRightWidth,
              borderBottomWidth: borderBottomWidth,
              borderLeftWidth: borderLeftWidth,
              borderColor: borderColor,
              opacity: Number.isFinite(opacity) ? opacity : 1,
              isDecorLine: isDecorLine,
              marginLeft: marginLeft,
              marginRight: marginRight,
            });
          }
        });
      });

      // 提取表格数据
      function extractTableData(table) {
        const rows = [];
        const rowHeights = [];
        const trs = table.querySelectorAll('tr');
        trs.forEach(tr => {
          const trRect = tr.getBoundingClientRect();
          const cells = [];
          const tds = tr.querySelectorAll('td, th');
          tds.forEach(td => {
            const tdStyle = window.getComputedStyle(td);
            cells.push({
              text: td.innerText || td.textContent || '',
              backgroundColor: tdStyle.backgroundColor,
              textAlign: tdStyle.textAlign,
              fontWeight: tdStyle.fontWeight,
              color: tdStyle.color,
              fontSize: tdStyle.fontSize,
              fontFamily: tdStyle.fontFamily,
            });
          });
          if (cells.length > 0) {
            rows.push(cells);
            rowHeights.push(trRect.height);
          }
        });
        return { rows, rowHeights };
      }

      // 提取富文本 runs（递归提取，样式继承合并）
      function extractRuns(element) {
        const runs = [];

        function mergeStyle(inherited, current) {
          if (!inherited) return current;
          return {
            fontWeight: current.fontWeight || inherited.fontWeight,
            fontStyle: current.fontStyle || inherited.fontStyle,
            // textDecoration 需要合并：父级和子级的装饰都要保留
            textDecoration: [inherited.textDecoration, current.textDecoration].filter(Boolean).join(' '),
            color: current.color !== inherited.color ? current.color : inherited.color,
            fontSize: current.fontSize || inherited.fontSize,
            fontFamily: current.fontFamily || inherited.fontFamily,
          };
        }

        function processNode(node, inheritedStyle) {
          if (node.nodeType === 3) {
            // 纯文本节点
            const text = node.textContent;
            if (text.trim()) {
              const style = inheritedStyle || window.getComputedStyle(element);
              // 清理文本：去除首尾空白，并清理内部的换行和多余空格
              const cleanedText = text.trim().replace(/\s+/g, ' ');
              runs.push({
                text: cleanedText,
                bold: style.fontWeight === 'bold' || parseInt(style.fontWeight) >= 700,
                italic: style.fontStyle === 'italic',
                underline: style.textDecoration.includes('underline'),
                color: style.color,
                fontSize: style.fontSize,
                fontFamily: style.fontFamily,
              });
            }
          } else if (node.nodeType === 1) {
            // 处理 <br> 标签：添加换行 run
            if (node.tagName === 'BR') {
              const style = inheritedStyle || window.getComputedStyle(element);
              runs.push({
                text: '',
                breakLine: true,
                bold: style.fontWeight === 'bold' || parseInt(style.fontWeight) >= 700,
                italic: style.fontStyle === 'italic',
                underline: style.textDecoration.includes('underline'),
                color: style.color,
                fontSize: style.fontSize,
                fontFamily: style.fontFamily,
              });
              return;
            }

            // 标签节点：获取该标签的样式，与继承样式合并
            const style = window.getComputedStyle(node);

            const tagStyle = {
              fontWeight: style.fontWeight,
              fontStyle: style.fontStyle,
              textDecoration: style.textDecoration,
              color: style.color,
              fontSize: style.fontSize,
              fontFamily: style.fontFamily,
            };

            // 合并继承样式
            const mergedStyle = mergeStyle(inheritedStyle, tagStyle);

            // 递归处理子节点
            node.childNodes.forEach(child => processNode(child, mergedStyle));
          }
        }

        element.childNodes.forEach(node => processNode(node, null));

        return runs.length > 0 ? runs : null;
      }

      // 提取表格
      const tables = [];
      document.querySelectorAll('table').forEach(table => {
        const rect = table.getBoundingClientRect();
        const style = window.getComputedStyle(table);
        const tableData = extractTableData(table);
        tables.push({
          x: rect.left,
          y: rect.top,
          width: rect.width,
          height: rect.height,
          rows: tableData.rows,
          rowHeights: tableData.rowHeights,
          backgroundColor: style.backgroundColor,
          borderColor: style.borderColor,
          borderWidth: parseFloat(style.borderWidth) || 0,
        });
      });

      return { elements: results, tables, svgs, pseudoElements, pageInfo, bodyBackgroundColor, debugInfo };
    });

    // 注意：不再进行后处理
    // Puppeteer提取的数据已经是准确的，任何修改都会破坏原始布局

    if (verbose) console.log(`[extract] 提取到 ${extracted.elements.length} 个元素, ${extracted.tables.length} 个表格`);
    if (verbose) console.log(`[extract] 页面尺寸: ${extracted.pageInfo.width}x${extracted.pageInfo.height}px`);

    // 输出调试信息 - 直接分析提取的元素
    if (verbose) {
      console.log('=== 调试信息：所有有背景色的元素位置 ===');
      const keyElements = extracted.elements.filter(el => 
        el.backgroundColor && el.backgroundColor !== 'rgba(0, 0, 0, 0)' && el.backgroundColor !== 'transparent'
      );
      console.log(`找到 ${keyElements.length} 个有背景色的元素`);
      
      keyElements.forEach((el, idx) => {
        const hasText = el.text && el.text.trim();
        const textPreview = hasText ? el.text.substring(0, 30) + '...' : '无文本';
        console.log(`元素 ${idx+1}: tag=${el.tag}, x=${el.x}, y=${el.y}, w=${el.width}, h=${el.height}, 背景色=${el.backgroundColor}, 文本="${textPreview}"`);
      });
      console.log('==========================');
      
      // 专门检查标题区域的 H1 和 P 元素
      console.log('=== 调试信息：标题区域 H1 和 P 元素 ===');
      const titleElements = extracted.elements.filter(el => 
        el.tag === 'H1' || el.tag === 'P'
      );
      console.log(`找到 ${titleElements.length} 个 H1/P 元素`);
      
      titleElements.forEach((el, idx) => {
        console.log(`元素 ${idx+1}: tag=${el.tag}, text="${el.text}", x=${el.x}, y=${el.y}, w=${el.width}, h=${el.height}`);
      });
      console.log('==========================');
      
      // 专门检查伪元素
      console.log('=== 调试信息：伪元素 ===');
      console.log(`找到 ${extracted.pseudoElements.length} 个伪元素`);
      extracted.pseudoElements.forEach((pe, idx) => {
        console.log(`伪元素 ${idx+1}: tag=${pe.tag}, text="${pe.text}", x=${pe.x}, y=${pe.y}, w=${pe.width}, h=${pe.height}, bg=${pe.backgroundColor}`);
      });
      console.log('==========================');
    }

    return {
      elements: extracted.elements,
      tables: extracted.tables,
      svgs: extracted.svgs,
      pseudoElements: extracted.pseudoElements,
      pageWidth: extracted.pageInfo.width,
      pageHeight: extracted.pageInfo.height,
      bodyBackgroundColor: extracted.bodyBackgroundColor,
    };

  } finally {
    await browser.close();
  }
}

module.exports = { extractFromHtml };
