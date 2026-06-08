# html2pptx - AI生成H5页面提示词

## 使用说明

将以下提示词复制给AI（如ChatGPT、Claude等），让它帮您生成可以完美转换成PPT的H5页面！

---

## 完整提示词（直接复制使用）

```
请帮我创建一个纯静态HTML+CSS页面，用于html2pptx工具转换为PPT。请严格遵守以下规范：

【页面尺寸规范】
- 页面固定宽度：1280px
- 页面固定高度：720px
- 使用16:9比例，无滚动条
- 所有内容必须在1280×720范围内完全可见

【技术限制】
- 禁止使用JavaScript
- 禁止使用外部CSS框架（如Tailwind、Bootstrap等）
- 禁止使用动画、过渡效果（animation, transition）
- 禁止使用 box-shadow（PPT 不支持阴影）
- 只使用内联CSS或&lt;style&gt;标签内的纯CSS

【支持的HTML元素】
- 文本类：h1, h2, h3, h4, h5, h6, p, span, a, li, b, strong, i, em, u
- 容器类：div, section, article, header, footer, main, nav, aside
- 媒体类：img（使用绝对URL）, svg
- 表格类：table, tr, td, th, thead, tbody, caption
- 装饰类：hr
- 列表类：ul, ol

【支持的CSS样式】
- 背景：background-color（纯色）、background-image（linear-gradient 渐变）
- 渐变：linear-gradient() 支持多色标 + 角度/方向 + 透明度，自动转为 PPT 原生渐变
- 边框：border, border-width, border-color, border-style, border-radius
- 独立边框：border-top/right/bottom/left-width/color/style（支持单边边框）
- 字体：font-family, font-size, font-weight, font-style, text-decoration
- 文本：color, text-align, line-height
- 布局：padding（四方向独立）, opacity, white-space, transform（优先使用 rotate；translate 会尽量解析，scale 不建议作为关键布局手段）
- 伪元素：::before, ::after（支持装饰条、图标、背景块等）  
  *注意：伪元素支持有限，对于关键视觉元素（如列表项圆点、图标），建议使用真实DOM元素以确保转换可靠性*
- 富文本：同一元素内支持混合加粗/斜体/下划线/颜色（如 &lt;b&gt;、&lt;i&gt;、&lt;u&gt;、&lt;span style="color:..."&gt;）

【字体规范】
- 使用安全字体：Microsoft YaHei（微软雅黑）, Arial, SimHei（黑体）, SimSun（宋体）
- 避免使用小众或网络字体

【颜色规范】
- 使用十六进制：#RRGGBB 或 #RGB
- 或使用RGB/RGBA：rgb(0,0,0), rgba(0,0,0,0.5)
- 透明度使用opacity属性或rgba的alpha通道

【图片规范】
- 使用绝对URL（https://...）
- 避免使用本地相对路径

【布局建议】
- 使用绝对定位或flex布局确保元素位置精确
- 避免使用JavaScript计算位置
- 关键内容留有适当边距，不要贴边

【文本与换行规范】
- 普通段落、说明文字允许自动换行
- 如果某行必须保持单行，例如短标签、列表项、金额、页码、编号标题，请显式设置 white-space: nowrap
- 如果需要保留 HTML 中的视觉换行，优先使用块级元素或 display:block，不要依赖多个空格模拟换行
- 如果需要“标签 + 正文”分两行显示，可以使用 &lt;span class="label"&gt;现状问题&lt;/span&gt;正文，并设置 .label { display: block; }

【文本框边距与对齐规范】
- 转换后 PPT 文本默认垂直居中，上下内部边距为 0
- 不要依赖 padding-top / padding-bottom 精细控制文字垂直位置
- 如需控制文字与边框的水平距离，可以使用 padding-left / padding-right
- 普通文本请明确设置 text-align: left/center/right
- 图标、圆形编号、徽章等需要居中时，建议使用 display:flex; justify-content:center; align-items:center
- 不要依赖背景色或边框来推断文字居中

【装饰符号规范】
- 列表项圆点、勾选符号、编号徽章建议使用独立 DOM 元素，例如 &lt;span class="li-dot"&gt;✓&lt;/span&gt; 或 &lt;span class="badge"&gt;②&lt;/span&gt;
- 装饰符号与正文建议使用 flex 布局：display:flex; align-items:center
- 装饰符号设置固定 width/height，并设置 margin-right，避免与正文重叠
- 不要把装饰符号直接写进正文文本中，避免 PPT 中重复或难以单独定位

【伪元素层级规范】
- 使用 ::before / ::after 做背景线、装饰线时，请明确设置 position 和 z-index
- 如果装饰线应位于内容下方，例如流程线，请设置装饰线 z-index 低于内容元素
- 关键内容不要完全依赖伪元素，优先使用真实 DOM

【内容页头部与页脚统一规范】（除封面页/目录页外的所有内容页）
- 头部（page-header）：所有内容页顶部样式必须保持一致
  · 高度固定：80px
  · 左侧：页面标题（h2，font-size: 28px，color: #1a2a4f，font-weight: bold）
  · 右侧：章节序号或章节名称（font-size: 16px，color: #6b7280）
  · 底部装饰条：border-bottom: 3px solid #1a2a4f（或主题色）
  · 内边距：padding: 20px 60px
- 页脚（page-footer）：所有内容页底部样式必须保持一致
  · 高度固定：50px
  · 左侧：公司名称或项目名称（font-size: 14px，color: #6b7280）
  · 右侧：页码（如 "03 / 16"，font-size: 14px，color: #6b7280）
  · 顶部装饰线：border-top: 1px solid #e5e7eb
  · 内边距：padding: 15px 60px
- 主体内容区域（page-body）：高度 = 720 - 80(头部) - 50(页脚) = 590px
- 所有内容页的头部标题字号、颜色、装饰线条样式必须保持一致
- 所有内容页的页脚字号、颜色、页码格式必须保持一致
- 推荐使用 flex 纵向布局：header(80px) + body(590px, flex:1) + footer(50px)

【转换优化提示】
- 有背景色的容器请明确设置background-color
- 渐变背景使用 linear-gradient()，支持多色标、角度、透明度，自动转为 PPT 原生渐变
- 文本元素请明确设置font-size和color
- 圆角使用border-radius属性
- 圆形元素请设置border-radius: 50%且宽高相等
- 装饰条可使用::before或::after伪元素实现，但关键视觉元素（列表项圆点、图标、编号徽章）建议使用真实DOM元素
- 单边边框（如下划线装饰）使用 border-bottom 而非 border
- 必须单行显示的文本设置 white-space: nowrap
- 标签和正文需要分行时，标签使用 display:block
- 避免使用 CSS 变量 var(--xxx)，直接写具体颜色值

现在请根据以下需求创建H5页面：
[在此处填写您的具体需求，例如：创建一个封面页，包含标题"项目汇报"，副标题"2024年度总结"，底部有公司logo和日期]
```

---

## 提示词模板（可按需修改）

### 模板1：封面页
```
请帮我创建一个纯静态HTML+CSS封面页，用于html2pptx转换。
要求：1280×720px，无JS，纯CSS。

页面内容：
- 顶部：主标题（大字号）
- 中部：副标题（中等字号）
- 底部：日期和汇报人

请使用简洁的商务风格，蓝白配色。
```

### 模板2：内容页（卡片布局）
```
请帮我创建一个纯静态HTML+CSS内容页，用于html2pptx转换。
要求：1280×720px，无JS，纯CSS。

页面结构（严格遵循统一头部和页脚规范）：
- 头部（80px高）：左侧页面标题，右侧章节序号，底部3px主题色装饰条
- 主体内容区（590px高）：4个信息卡片（2×2布局）
  每个卡片包含：图标占位+标题+简短说明
- 页脚（50px高）：左侧项目名称，右侧页码（如 "03 / 16"），顶部1px灰色分隔线

请使用现代简洁风格，头部和页脚样式需与其他内容页保持一致。
```

### 模板3：数据展示页
```
请帮我创建一个纯静态HTML+CSS数据页，用于html2pptx转换。
要求：1280×720px，无JS，纯CSS。

页面结构（严格遵循统一头部和页脚规范）：
- 头部（80px高）：左侧页面标题，右侧章节序号，底部3px主题色装饰条
- 主体内容区（590px高）：左侧表格（5列3行），右侧文字说明区域
- 页脚（50px高）：左侧项目名称，右侧页码，顶部1px灰色分隔线

请使用清晰的数据展示风格，头部和页脚样式需与其他内容页保持一致。
```

---

## 最佳实践示例

### 示例1：最简单的有效H5
```html
&lt;!DOCTYPE html&gt;
&lt;html&gt;
&lt;head&gt;
    &lt;meta charset="UTF-8"&gt;
    &lt;title&gt;PPT Slide&lt;/title&gt;
    &lt;style&gt;
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            width: 1280px;
            height: 720px;
            background: #ffffff;
            font-family: "Microsoft YaHei", Arial;
        }
        .slide {
            width: 1280px;
            height: 720px;
            padding: 60px;
            background: #667eea;
        }
        h1 {
            color: #ffffff;
            font-size: 72px;
            margin-bottom: 30px;
        }
        p {
            color: #f0f0f0;
            font-size: 32px;
        }
    &lt;/style&gt;
&lt;/head&gt;
&lt;body&gt;
    &lt;div class="slide"&gt;
        &lt;h1&gt;欢迎使用 html2pptx&lt;/h1&gt;
        &lt;p&gt;纯静态HTML+CSS → 可编辑PPT&lt;/p&gt;
    &lt;/div&gt;
&lt;/body&gt;
&lt;/html&gt;
```

### 示例2：内容页（带统一头部和页脚）
```html
&lt;!DOCTYPE html&gt;
&lt;html&gt;
&lt;head&gt;
    &lt;meta charset="UTF-8"&gt;
    &lt;title&gt;内容页&lt;/title&gt;
    &lt;style&gt;
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 1280px;
            height: 720px;
            background: #f5f7fa;
            font-family: "Microsoft YaHei";
        }
        .slide {
            width: 1280px;
            height: 720px;
            display: flex;
            flex-direction: column;
        }
        /* 统一头部样式 */
        .page-header {
            height: 80px;
            padding: 20px 60px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 3px solid #1a2a4f;
            background: #ffffff;
        }
        .page-header .page-title {
            font-size: 28px;
            color: #1a2a4f;
            font-weight: bold;
        }
        .page-header .section-info {
            font-size: 16px;
            color: #6b7280;
        }
        /* 主体内容区 */
        .page-body {
            flex: 1;
            height: 590px;
            padding: 40px 60px;
        }
        .cards {
            display: flex;
            gap: 30px;
            height: 100%;
        }
        .card {
            flex: 1;
            background: #fff;
            padding: 30px;
            border-radius: 12px;
            border: 1px solid #e0e0e0;
        }
        .card h3 {
            font-size: 28px;
            color: #1a2a4f;
            margin-bottom: 15px;
        }
        .card p {
            font-size: 18px;
            color: #666;
            line-height: 1.6;
        }
        /* 统一页脚样式 */
        .page-footer {
            height: 50px;
            padding: 15px 60px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-top: 1px solid #e5e7eb;
            background: #ffffff;
        }
        .page-footer .project-name {
            font-size: 14px;
            color: #6b7280;
        }
        .page-footer .page-number {
            font-size: 14px;
            color: #6b7280;
        }
    &lt;/style&gt;
&lt;/head&gt;
&lt;body&gt;
    &lt;div class="slide"&gt;
        &lt;!-- 统一头部 --&gt;
        &lt;div class="page-header"&gt;
            &lt;div class="page-title"&gt;核心功能&lt;/div&gt;
            &lt;div class="section-info"&gt;第二章 · 解决方案&lt;/div&gt;
        &lt;/div&gt;
        &lt;!-- 主体内容 --&gt;
        &lt;div class="page-body"&gt;
            &lt;div class="cards"&gt;
                &lt;div class="card"&gt;
                    &lt;h3&gt;元素提取&lt;/h3&gt;
                    &lt;p&gt;精确解析HTML元素位置与样式&lt;/p&gt;
                &lt;/div&gt;
                &lt;div class="card"&gt;
                    &lt;h3&gt;样式转换&lt;/h3&gt;
                    &lt;p&gt;完美还原CSS样式到PPT&lt;/p&gt;
                &lt;/div&gt;
                &lt;div class="card"&gt;
                    &lt;h3&gt;批量处理&lt;/h3&gt;
                    &lt;p&gt;支持多文件合并转换&lt;/p&gt;
                &lt;/div&gt;
            &lt;/div&gt;
        &lt;/div&gt;
        &lt;!-- 统一页脚 --&gt;
        &lt;div class="page-footer"&gt;
            &lt;div class="project-name"&gt;html2pptx 项目汇报&lt;/div&gt;
            &lt;div class="page-number"&gt;03 / 16&lt;/div&gt;
        &lt;/div&gt;
    &lt;/div&gt;
&lt;/body&gt;
&lt;/html&gt;
```

---

## 常见错误及避免方法

| 错误做法 | 正确做法 |
|---------|---------|
| 使用 box-shadow | 用 border 或伪元素模拟阴影效果 |
| 使用JavaScript | 纯CSS实现所有效果 |
| 页面尺寸不固定 | 严格1280×720px |
| 使用Tailwind/Bootstrap | 手写纯CSS |
| 相对路径图片 | 使用绝对URL图片 |
| 小众字体 | 使用微软雅黑/Arial |
| 动画/过渡效果 | 静态设计 |
| 使用 CSS 变量 var(--xxx) | 直接写具体颜色值 |
| 使用复杂伪元素装饰 | 关键视觉元素（列表项圆点、图标、编号徽章）使用真实DOM元素 |
| 把装饰符号直接写入正文 | 用独立 DOM 元素承载符号，并给正文留出 margin-right/偏移空间 |
| 依赖多个空格制造换行 | 使用 display:block 或明确的块级结构 |
| 必须单行的文本未设置 nowrap | 添加 white-space: nowrap |

---

## 推荐设计风格

### 风格1：商务简洁风
- 配色：深蓝、浅灰、白色
- 字体：微软雅黑
- 特点：线条简洁，专业大方

### 风格2：现代科技风
- 配色：深蓝、黑、白，可搭配 linear-gradient 渐变背景
- 字体：Arial + 微软雅黑
- 特点：几何图形，留白充足，暗色主题

### 风格3：清新活力风
- 配色：浅绿、浅蓝、米白
- 字体：微软雅黑
- 特点：圆角元素，色彩明快

---

## 自检清单

生成H5后，请检查：

- [ ] 页面尺寸：1280×720px
- [ ] 无JavaScript代码
- [ ] 无外部CSS框架
- [ ] 无 box-shadow（用 border 或伪元素替代）
- [ ] 无动画/过渡效果
- [ ] 无 CSS 变量 var(--xxx)
- [ ] 使用安全字体（微软雅黑、Arial 等）
- [ ] 图片使用绝对URL
- [ ] 所有内容在可视范围内
- [ ] 颜色使用十六进制或RGB/RGBA
- [ ] 内容页头部统一：高度80px，标题字号/颜色/装饰条一致
- [ ] 内容页页脚统一：高度50px，项目名称/页码字号/颜色一致
- [ ] 主体内容区高度为 590px（720 - 80 - 50）
- [ ] 关键装饰元素（列表项圆点、图标、编号徽章）使用真实DOM元素而非伪元素
- [ ] 装饰符号设置固定宽高和 margin-right，避免与正文重叠
- [ ] 必须保持单行的文本已设置 white-space: nowrap
- [ ] 需要保留视觉换行的标签已使用 display:block
- [ ] 图标、圆形编号、徽章等居中元素已使用 flex 居中
- [ ] 伪元素装饰线已明确 position 和 z-index

检查通过后，即可使用html2pptx转换！
