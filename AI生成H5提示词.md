# html2pptx - AI生成H5页面提示词

## 📋 使用说明

将以下提示词复制给AI（如ChatGPT、Claude等），让它帮您生成可以完美转换成PPT的H5页面！

---

## 🎯 完整提示词（直接复制使用）

```
请帮我创建一个纯静态HTML+CSS页面，用于html2pptx工具转换为PPT。请严格遵守以下规范：

【页面尺寸规范】
- 页面固定宽度：1280px
- 页面固定高度：720px
- 使用16:9比例，无滚动条
- 所有内容必须在1280×720范围内完全可见

【技术限制】
- ❌ 禁止使用JavaScript
- ❌ 禁止使用外部CSS框架（如Tailwind、Bootstrap等）
- ❌ 禁止使用动画、过渡效果（animation, transition）
- ❌ 禁止使用 box-shadow（PPT 不支持阴影）
- ✅ 只使用内联CSS或<style>标签内的纯CSS

【支持的HTML元素】
- 文本类：h1, h2, h3, h4, h5, h6, p, span, a, li, b, strong, i, em, u
- 容器类：div, section, article, header, footer, main, nav, aside
- 媒体类：img（使用绝对URL）, svg
- 表格类：table, tr, td, th, thead, tbody, caption
- 装饰类：hr
- 列表类：ul, ol

【支持的CSS样式】
✅ 背景：background-color（纯色）、background-image（linear-gradient 渐变）
✅ 渐变：linear-gradient() 支持多色标 + 角度/方向 + 透明度，自动转为 PPT 原生渐变
✅ 边框：border, border-width, border-color, border-style, border-radius
✅ 独立边框：border-top/right/bottom/left-width/color/style（支持单边边框）
✅ 字体：font-family, font-size, font-weight, font-style, text-decoration
✅ 文本：color, text-align, line-height
✅ 布局：padding（四方向独立）, opacity, transform（rotate/translate/scale）
✅ 伪元素：::before, ::after（支持装饰条、图标、背景块等）
✅ 富文本：同一元素内支持混合加粗/斜体/下划线/颜色（如 <b>、<i>、<u>、<span style="color:...">）

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

【转换优化提示】
- 有背景色的容器请明确设置background-color
- 渐变背景使用 linear-gradient()，支持多色标、角度、透明度，自动转为 PPT 原生渐变
- 文本元素请明确设置font-size和color
- 圆角使用border-radius属性
- 圆形元素请设置border-radius: 50%且宽高相等
- 装饰条使用::before或::after伪元素实现
- 单边边框（如下划线装饰）使用 border-bottom 而非 border
- 避免使用 CSS 变量 var(--xxx)，直接写具体颜色值

现在请根据以下需求创建H5页面：
[在此处填写您的具体需求，例如：创建一个封面页，包含标题"项目汇报"，副标题"2024年度总结"，底部有公司logo和日期]
```

---

## 📝 提示词模板（可按需修改）

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

页面内容：
- 顶部：页面标题
- 主体：4个信息卡片（2×2布局）
  每个卡片包含：图标占位+标题+简短说明
- 底部：页码

请使用现代简洁风格。
```

### 模板3：数据展示页
```
请帮我创建一个纯静态HTML+CSS数据页，用于html2pptx转换。
要求：1280×720px，无JS，纯CSS。

页面内容：
- 顶部：标题
- 左侧：表格（5列3行）
- 右侧：文字说明区域

请使用清晰的数据展示风格。
```

---

## ✅ 最佳实践示例

### 示例1：最简单的有效H5
```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>PPT Slide</title>
    <style>
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
    </style>
</head>
<body>
    <div class="slide">
        <h1>欢迎使用 html2pptx</h1>
        <p>纯静态HTML+CSS → 可编辑PPT</p>
    </div>
</body>
</html>
```

### 示例2：带卡片的页面
```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>内容页</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 1280px;
            height: 720px;
            background: #f5f7fa;
            font-family: "Microsoft YaHei";
        }
        .container {
            width: 1280px;
            height: 720px;
            padding: 50px;
        }
        .title {
            font-size: 48px;
            color: #333;
            margin-bottom: 40px;
        }
        .cards {
            display: flex;
            gap: 30px;
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
            color: #333;
            margin-bottom: 15px;
        }
        .card p {
            font-size: 18px;
            color: #666;
            line-height: 1.6;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2 class="title">核心功能</h2>
        <div class="cards">
            <div class="card">
                <h3>元素提取</h3>
                <p>精确解析HTML元素位置与样式</p>
            </div>
            <div class="card">
                <h3>样式转换</h3>
                <p>完美还原CSS样式到PPT</p>
            </div>
            <div class="card">
                <h3>批量处理</h3>
                <p>支持多文件合并转换</p>
            </div>
        </div>
    </div>
</body>
</html>
```

---

## ⚠️ 常见错误及避免方法

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

---

## 🎨 推荐设计风格

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

## 🔍 自检清单

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

检查通过后，即可使用html2pptx转换！
