# HTML Deck Export Skill

![Skill](https://img.shields.io/badge/Skill-Agent-111111?style=flat-square)
![Codex](https://img.shields.io/badge/Codex-Supported-222222?style=flat-square)
![HTML Deck](https://img.shields.io/badge/HTML-Deck-0A7CFF?style=flat-square)
![PDF](https://img.shields.io/badge/Export-PDF-FF7A1A?style=flat-square)
![PPTX](https://img.shields.io/badge/Export-PPTX-00AEEF?style=flat-square)

一个给 Codex / Claude Code / 本地 Agent 使用的 HTML 演示稿导出 skill。

它的目标很直接：把已经做好的本地 HTML 展示稿、路演预览稿、招商画册式 web PPT，导出成适合交付和编辑的文件：

- 高保真 16:9 PDF
- 16:9 PPTX
- PPTX 中保留可编辑文字框
- 同步拆出底图、蒙版、完整截图和文字坐标 JSON

这个 skill 来自一次真实的“招商画册 / 路演预览稿”交付流程，重点不是从零生成 PPT，而是把已经排好的 HTML deck 稳定转成 PDF 和 PPTX。

## 30 秒开始

```bash
git clone https://github.com/summer817-AI/summer-html-deck-export.git
```

安装到 Codex skill 目录：

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.codex\skills" | Out-Null
Copy-Item -Recurse .\summer-html-deck-export "$env:USERPROFILE\.codex\skills\html-deck-export"
```

安装依赖：

```powershell
pip install beautifulsoup4 pillow python-pptx pypdf playwright
python -m playwright install chromium
```

安装后直接对 Codex 说：

```text
请使用 html-deck-export，把这个 HTML 路演预览稿导出成 16:9 PDF 和可编辑文字 PPTX，并保留底图、蒙版和图层素材。
```

## 它能做什么

### 1. HTML 转高保真 PDF

使用 Playwright 调用浏览器渲染 HTML，再导出 16:9 PDF。这个 PDF 是视觉交付的主版本，适合给客户、领导或团队预览。

### 2. HTML 转可编辑 PPTX

PPTX 不是简单把每页压成一张图，而是按下面的结构生成：

1. 背景图层
2. 半透明蒙版图层
3. 顶层可编辑文字框

这样可以在 PowerPoint 里继续改标题、正文、页码、标注等文字。

### 3. 拆分图层素材

导出目录会保留完整过程文件，方便人工复查或后续重组：

- `layers/background/`: 每页裁切后的底图
- `layers/mask/`: 每页蒙版
- `layers/full/`: 每页完整截图
- `layers/source-background/`: 原始底图副本
- `editable-text-boxes.json`: 浏览器抽取出的文字框坐标和样式
- `hybrid-export-report.txt`: 导出报告

## 适合 / 不适合

适合：

- HTML/CSS 已经排好版，需要转成 PDF 或 PPTX 交付
- 路演稿、招商画册、方案汇报、电子画册式 deck
- 每页有大图底图、蒙版、标题和信息块
- 需要 PDF 保真，同时又希望 PPTX 文字还能编辑

不适合：

- 还没有 HTML，只是想从大纲直接生成完整 PPT
- 复杂动画、视频、交互组件必须在 PPTX 中完整保留
- 需要 100% 还原 CSS 排版为 PowerPoint 原生形状
- 大量表格、复杂 SVG、canvas 图表都要转成可编辑对象

## 输入要求

默认脚本适配这样的 HTML deck：

- 每页是 `section.slide`
- 每页底图是 `img.slide-bg`
- 页面蒙版通常由 `.wash` 或类似 CSS 控制
- 设计比例为 16:9

如果你的 HTML 结构不同，PDF 通常仍然可以正常导出；PPTX 的可编辑文字抽取可能需要修改 `scripts/html_deck_hybrid_export.py` 里的文本选择器。

## 命令行调用

安装后可以直接运行脚本：

```powershell
python "$env:USERPROFILE\.codex\skills\html-deck-export\scripts\html_deck_hybrid_export.py" `
  "D:\path\to\deck.html" `
  -o "D:\path\to\export-output"
```

常用参数：

```powershell
python .\scripts\html_deck_hybrid_export.py `
  ".\index.html" `
  -o ".\html-hybrid-export" `
  --width 1600 `
  --height 900 `
  --timeout 30000
```

如果 Chrome / Edge 不在默认路径，可以手动指定：

```powershell
python .\scripts\html_deck_hybrid_export.py `
  ".\index.html" `
  -o ".\html-hybrid-export" `
  --chrome "C:\Program Files\Google\Chrome\Application\chrome.exe"
```

## 输出文件

导出完成后，输出目录通常包含：

```text
html-deck-16x9-playwright.pdf
html-deck-16x9-editable-text.pptx
editable-text-boxes.json
hybrid-export-report.txt
layers/
  background/
  mask/
  full/
  source-background/
```

建议检查顺序：

1. 先看 `hybrid-export-report.txt`，确认页数、PDF 页数、PPTX 路径和文字框数量。
2. 再看 `html-deck-16x9-playwright.pdf`，它是视觉保真版本。
3. 最后打开 `html-deck-16x9-editable-text.pptx`，检查可编辑文字是否覆盖完整。

## Agent 调用示例

把下面任意一条发给 Codex / Claude Code：

```text
请用 html-deck-export 把这个 HTML 展示稿转成 PDF 和 PPTX，输出到同级 export 文件夹。
```

```text
请使用 html-deck-export。输入文件是 D:\project\deck\index.html，输出到 D:\project\deck\html-hybrid-export，生成 PDF、PPTX 和图层素材。
```

```text
请用 html-deck-export 检查这个 HTML deck 是否符合导出结构，如果符合就导出；如果不符合，先告诉我需要调整哪些 DOM selector。
```

## 质量原则

- PDF 是视觉基准。若 PDF 正常而 PPTX 有轻微偏差，优先以 PDF 作为展示版本。
- PPTX 是可编辑工作稿。它尽量保留文字框，但不承诺把所有 CSS 效果转成 PowerPoint 原生对象。
- 如果 PPTX 丢文字，优先扩展 `extract_text_boxes()` 里的 selector。
- 如果字体大小偏差明显，检查 `editable-text-boxes.json`，再调整 `make_editable_pptx()` 中的字号缩放。
- 如果页面比例错乱，先确认 HTML 是否按 16:9 固定布局，并避免打印样式覆盖屏幕样式。

## 依赖

Python 包：

```bash
pip install beautifulsoup4 pillow python-pptx pypdf playwright
python -m playwright install chromium
```

浏览器：

- Google Chrome
- Microsoft Edge
- Playwright Chromium

脚本会优先查找常见 Windows Chrome / Edge 路径，也可以用 `--chrome` 手动传入浏览器路径。

## 文件结构

```text
html-deck-export/
  SKILL.md
  README.md
  agents/
    openai.yaml
  scripts/
    html_deck_hybrid_export.py
```

`SKILL.md` 是 Agent 读取的核心说明；`README.md` 面向人类读者；`scripts/html_deck_hybrid_export.py` 是实际导出脚本。

## License

当前仓库未单独声明许可证。使用、复制或二次分发前，请先确认仓库所有者的授权边界。
