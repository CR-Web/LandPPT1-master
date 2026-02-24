# 重构 PPT 导入与生成引擎：实现“像素级”保真与智能视觉适配

针对 `business_blue_01.pptx` 导入效果不佳及生成 PPT 不美观的问题，我将执行以下大规模重构：

## 1. 升级 PPTX 导入引擎 (Fidelity Upgrade)
- **文件**：[global_master_template_api.py](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/api/global_master_template_api.py)
- **改进内容**：
    - **SVG 矢量渲染**：重构 `_pptx_to_multislide_html`，不再将所有形状简单映射为 CSS 矩形。引入 SVG 渲染逻辑，支持更多 PPT 预设形状（如各种三角形、梯形、箭头等）。
    - **高精度 Z-Index**：优化层级计算，确保背景装饰元素（如您模板中的蓝色多边形）始终位于文字下方，且遮挡关系与 PPT 原始设计一致。
    - **背景资产提取**：改进对幻灯片母版背景的解析，确保导入后的 HTML 包含完整的渐变、图片和矢量路径。

## 2. 实现封面页“像素级”还原模式 (Strict Cover Mode)
- **文件**：[enhanced_ppt_service.py](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/services/enhanced_ppt_service.py)
- **改进内容**：
    - 在 `_generate_slide_with_template` 中为第 1 页开启“严格模式”。
    - **禁止 LLM 调整**：仅执行占位符替换，保留原始 HTML 结构、内联样式和所有 DOM 节点，确保封面与导入的模板 100% 相同。

## 3. 构建智能视觉适配引擎 (Adaptive UI Engine v2)
- **文件**：[enhanced_ppt_service.py](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/services/enhanced_ppt_service.py)、[design_prompts.py](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/services/prompts/design_prompts.py)
- **改进内容**：
    - **内容量感知布局**：根据 `content_points` 数量自动切换布局。例如：超过 5 条要点时自动启用双列布局，防止底部溢出。
    - **动态样式微调**：授权 LLM 动态修改 `font-size`（使用 `clamp()`）、`line-height`、`padding` 和 `gap`，使文字填充页面更加饱满美观。
    - **位置优化指令**：在 Prompt 中明确要求 LLM 调整绝对定位坐标（`top/left`），以修复元素重叠并达到视觉平衡。
    - **色彩和谐检查**：确保动态生成的组件颜色与模板设计基因（Design Genes）高度契合。

## 4. 强制画布与高度对齐
- **文件**：[enhanced_ppt_service.py](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/services/enhanced_ppt_service.py)
- **改进内容**：
    - 增强 `_ensure_fixed_canvas_size`，在每一页生成的最末端注入强制性的 1280x720 CSS 容器，确保所有页面的高度在导出时完全一致。

## 5. 前端预览同步
- **文件**：[global_master_templates.js](file:///e:/workandstudy/LandPPT-master/LandPPT-master/src/landppt/web/static/js/global_master_templates.js)
- **改进内容**：
    - 同步更新前端预览逻辑，使其与后端重构后的 1280x720 画布逻辑完全一致。

您是否同意按照此方案进行重构？确认后我将开始执行。