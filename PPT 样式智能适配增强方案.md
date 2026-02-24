# PPT 样式智能适配增强方案

## 核心目标
1. **封面页一致性**：强制封面页（第1页）严格保持模板原样，仅替换内容，不进行结构重排。
2. **样式动态适应**：除封面页外，所有页面根据内容量自动调整字号、间距和布局，但**必须**保留模板的视觉风格（背景、配色、装饰元素）。
3. **内容零溢出**：确保文字内容自动匹配容器大小，不超出边界，不遮挡装饰。

## 技术实现路径

### 1. 封面页锁定逻辑优化
**修改点**：在 `_generate_slide_with_template` 中，针对 `page_number == 1` 或 `slide_type == "title"` 的情况，跳过 LLM 布局重排。
- **当前逻辑**：已有部分判断，但需要加强。确保 LLM 不介入封面页的布局调整，仅进行文本填充。
- **增强措施**：在 `_apply_llm_template_adaptation` 的调用前增加强制守卫，封面页直接返回 `_process_template_html` 的结果（经过 sanitize 处理）。

### 2. 智能样式适配 (LLM Template Adaptation)
**修改点**：增强 `_build_template_adaptation_prompt` 提示词，强化“样式守恒”指令。
- **Prompt 优化**：
    - 明确要求“保留所有背景图片、SVG 装饰和绝对定位的视觉元素”。
    - 增加“文字容器自适应”指令：允许修改 `width/height/top/left` 但禁止覆盖装饰层。
    - 强制“字号动态计算”：根据字符数反比设定 `font-size`（如 >100字用 14px，<50字用 24px）。
    - 强调“配色继承”：必须使用模板原有的 CSS 变量或颜色值，严禁创造新颜色。

### 3. 内容填充与防溢出 (Process Template HTML)
**修改点**：在 `_process_template_html` 及其辅助函数中增强压缩逻辑。
- **智能截断与续页**：当前已有 `_apply_content_compression`，需要调整其阈值，使其更敏感。当内容过多时，优先缩小字号，达到极限（如 12px）后触发溢出机制（生成续页）。
- **DOM 结构清理**：在 `_semantic_fill_imported_slide_html` 后，增加一步“空元素清理”，移除未被填充的占位符容器，避免占据视觉空间。

### 4. 视觉一致性校验
**修改点**：在 LLM 返回结果后，增加 CSS 校验。
- 检查 `background` 属性是否被意外移除。
- 确认关键装饰元素（如 `.decoration`, `.bg-shape`）是否依然存在。

## 执行计划
1.  **修改 `enhanced_ppt_service.py`**：
    - 更新 `_generate_slide_with_template`：强化封面页跳过逻辑。
    - 更新 `_build_template_adaptation_prompt`：注入“样式守恒”和“动态适配”的强指令。
    - 微调 `_process_template_html`：优化初始填充的字号计算逻辑。
2.  **验证**：
    - 生成一份包含长文本和短文本的 PPT，检查封面页是否原样保留。
    - 检查内容页是否根据文字量自动调整了布局且未破坏背景。

## 用户确认
请确认是否同意上述方案，特别是关于“封面页严格锁定”和“内容页样式动态调整”的策略。确认后将开始修改代码。