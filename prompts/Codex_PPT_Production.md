# Codex PPT Production提示词

请使用当前Academic PPT Workflow仓库作为唯一生产权威，执行指定项目。用户只提供项目根目录，并可选择四个route与三个quality。

## 用户输入

```text
project: <PPT_WORKSPACE_HOME>/projects/<project_id>/
route: generate | enhance | template-fill | template-create
quality: quick | validated | full
```

从项目内读取`brief/presentation_brief.yaml`。只有当`brief/visual_brief.yaml`存在、通过schema验证且带有明确Human批准时才采用；`visual_brief.DRAFT.yaml`不得生效。不要要求用户提供内部规划对象、缓存标识或操作清单。

## 生产边界

- 事实、数字、图表和引文只来自注册输入来源；缺失内容标记`INFORMATION_REQUIRED`。
- 视觉brief只控制艺术方向，不得成为科学事实来源。
- 非数据型资产只有在`approved_assets/asset_registry.yaml`登记、Human批准且`scientific_evidence: false`时才能使用。
- 患者影像、ECG、超声、病理、显微图和真实科研图必须作为有来源的科学输入处理，不能由GPT概念资产替代。
- 保持native PptxGenJS默认后端，不接入外部PPT Skill。

## 视觉解析顺序

按以下优先级解析视觉要求：

```text
用户明确指令 > 已批准visual_brief > 获授权模板 > 既有Deck风格 > Canonical Workspace > 自动选择
```

例外不是可选项：

- `template-fill`必须保护正式模板的Master、Layout、placeholder、Logo、主题字体、主题颜色、页脚、页码和保留动画；品牌强制项高于visual brief。
- 小范围`enhance`且未明确要求重大重设计时，直接继承既有Deck风格，不要求GPT视觉方案，不重建未变化页面。

## 路线

- `generate`：从项目输入生成完整Deck；重大视觉设计使用已批准视觉brief或安全回退。
- `enhance`：仅规划、生成、替换和验证受影响页面；保留原Deck。
- `template-fill`：使用模板原生Master、Layout和placeholder填充，不以整页自由文本框覆盖模板。
- `template-create`：仅从获授权参考提取可复用风格规则，不复制敏感正文或未授权资产。

## 质量与输出

- `quick`：生成必要Draft和受影响页面QA，避免重复审计。
- `validated`：至少生成PPTX、PDF、contact sheet和最小科学/视觉审计。
- `full`：运行完整科学、视觉、PowerPoint及配置要求的交付验证。

每次使用独立run目录：

```text
output/<run_id>/draft/
output/<run_id>/final/
output/<run_id>/preview/
audit/<run_id>/
```

不得覆盖输入或既有正式输出。完成自动QA后，明确列出需要Human处理的科学不确定性和最终视觉审核项；Human批准前不得声称最终科学批准。
