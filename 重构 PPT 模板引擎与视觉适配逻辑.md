# 重构 PPT 模板引擎：实现像素级封面还原与智能视觉适配

## 核心目标
1. **封面 100% 还原**：确保生成的封面页与上传的 PPTX 模板完全一致，不产生任何结构性偏差。
2. **高度一致性**：所有页面强制锁定 1280x720 分辨率，确保导出 PDF/PPT 时高度对齐。
3. **智能视觉适配**：根据内容多少、位置、颜色自动调整样式，解决重叠、溢出、排版不美观等问题。

## 技术实施方案

### 1. 封面页“像素级”替换 (Strict Replacement Mode)
- **现状分析**：目前封面页虽然跳过了 LLM 适配，但在 `_process_template_html` 环节仍可能因 HTML 解析和重新构建导致样式丢失。
- **重构方案**：在 `enhanced_ppt_service.py` 中为第 1 页引入“严格替换”模式。
    - 使用 `BeautifulSoup` 精确锁定模板中的文本节点。
    - 仅替换 `{title}`, `{subtitle}` 等占位符，严禁触动任何 CSS 样式、SVG 形状或 DOM 结构。
    - 确保封面背景资产（图片、形状、渐变）100% 继承。

### 2. 强制画布与高度锁定
- **重构方案**：增强 `_ensure_fixed_canvas_size` 方法。
    - 在生成的 HTML 最外层注入一个 `1280x720` 的严格容器。
    - 注入全局 CSS 重置样式：`margin: 0; padding: 0; box-sizing: border-box; overflow: hidden;`。
    - 处理绝对定位溢出：如果内容坐标超出 720px，通过 CSS `transform: scale()` 进行整体缩放以适应画布。

### 3. 智能视觉适配引擎 (Adaptive UI Engine)
- **重构方案**：重写 `_apply_llm_template_adaptation` 及其相关 Prompt。
    - **样式守恒调整**：允许 LLM 在不改变设计风格的前提下，动态修改 `font-size`（如内容多则缩小）、`line-height`、`padding` 和 `gap`。
    - **位置优化**：授权 LLM 调整元素的绝对定位坐标（`top/left`），确保在内容增减时页面视觉平衡。
    - **重叠检测与修复**：在 Prompt 中增加“禁止重叠”指令，要求 LLM 检查并修复卡片、文字框之间的物理冲突。
    - **色板约束**：提取模板的 `Design Genes`（色板、字体库）并注入 LLM，确保动态生成的颜色不脱离模板范围。

### 4. 内容驱动的布局转换
- **重构方案**：改进 `_verify_layout` 逻辑。
    - 当内容点（Bullet Points）超过 5 个或文字量超过 500 字时，自动强制切换为“分栏”或“网格”布局，防止单列排版导致的页面溢出。

## 关键修改文件
- [enhanced_ppt_service.py](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/services/enhanced_ppt_service.py)：核心生成逻辑重构。
- [design_prompts.py](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/services/prompts/design_prompts.py)：智能适配 Prompt 升级。
- [global_master_templates.js](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/web/static/js/global_master_templates.js)：前端预览尺寸同步。

您是否同意按照此方案进行大规模代码重构？确认后我将开始执行。