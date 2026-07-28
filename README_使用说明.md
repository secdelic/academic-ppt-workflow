# 学术 PPT 自动制作工作流

这是一个本地、可重复、可审计、人工最终批准的学术 PowerPoint 工作流。
它不会把缺失信息“补写”成正式内容，也不会把自动 QA 等同于科学批准。

## 1. 最快开始

1. 将原始材料复制到对应的只读输入目录：

   - 文档：`input/documents/`
   - 数据：`input/data/`
   - 图片：`input/figures/`
   - 参考文献：`input/references/`
   - 参考 PPT 或风格文件：`input/style_reference/`

2. 编辑 `brief/presentation_brief.yaml`。该文件采用 JSON-compatible YAML，
   因此既是合法 YAML，也可在没有 PyYAML 时由 Python 标准库读取。

3. 运行：

   ```powershell
   python run_ppt_workflow.py --brief brief/presentation_brief.yaml
   ```

   如果 Node.js 不在 PATH 中：

   ```powershell
   python run_ppt_workflow.py `
     --brief brief/presentation_brief.yaml `
     --node "C:\path\to\node.exe" `
     --node-modules "C:\path\to\node_modules"
   ```

4. 在 `output/<project_name>/<timestamp>/` 中审核结果。

## 2. 支持的输入

工作流支持 DOCX、PDF、PPTX、XLSX、CSV、TSV、Markdown、TXT、PNG、
JPG/JPEG 和 SVG。

- PDF 先使用原生文本提取。
- 只有当 PDF 被确认是扫描件、原生提取失败且任务确实需要时，才允许另行
  启用 OCR。当前核心不自动运行 OCR。
- PNG/JPEG 只登记尺寸、格式和哈希，不静默 OCR。
- 所有文件写入 `audit/source_manifest.csv`，包括 SHA-256、大小、时间、
  解析状态、页数/工作表数、警告和敏感信息提示。

## 3. 如何填写 Presentation Brief

必填的科学/演示字段包括项目名、类型、目标、受众、语言、时长、目标页数、
核心信息和引用风格。没有答案时保留 `INFORMATION_REQUIRED`；工作流会阻断
或降级，而不是猜测。

`presentation_type` 可选：

- `research_report`
- `thesis_defense`
- `journal_club`
- `conference_talk`
- `project_update`
- `clinical_research_protocol`
- `bioinformatics_study`

默认结构在 `config/profiles.yaml` 中，可以按项目复制并修改，但不要把项目
事实写成全局默认。

## 4. 证据登记约定

普通文本会进入内容清单，但不会自动变成高置信度正式结论。若希望确定性测试
或受控项目显式注册候选结论，可在源文件中使用：

```text
CLAIM|claim_text|source_location|evidence_type|confidence|allowed_wording|prohibited_overstatement
```

这不是绕过人工审核的机制。每个 `CLAIM` 仍会进入证据表、来源脚注和人工审核。

工作流生成：

- `staging/content_inventory.md`
- `staging/evidence_inventory.csv`
- `staging/figure_inventory.csv`
- `staging/table_inventory.csv`
- `audit/unresolved_items.md`

## 5. 故事线和页面结构

PPTX 生成前必须存在：

- `staging/deck_outline.md`
- `staging/storyboard.csv`

每页只有一个主要目的和一个关键消息。来源不足时使用中性标题或
`INFORMATION_REQUIRED`，不会生成夸大的结论性标题。

页面数量以 brief、类型 profile 和可用证据共同控制。内容过多时应拆页，不能
通过无限缩小字号塞入页面。

## 6. 参考 PPT 和设计风格

将参考文件放入 `input/style_reference/`，并在 brief 的
`style_reference_files` 中列出相对路径。当前 v1 会登记并保护这些文件，
但不会自动复制其母版或版式；在启用模板继承前，必须先人工确认模板版权、
母版结构和允许复用的页面。未启用时使用 `restrained_biomedical` 主题。

主题颜色、字体回退、字号和网格位于 `config/design_system.yaml`。只使用本机
已安装字体，不复制、嵌入或分发字体文件。

## 7. 生成后端和依赖

- 默认后端：PptxGenJS 4.0.1（MIT）。
- 受控内部验证后端：`--backend artifact_tool`。

Artifact Tool 在本环境中的许可证仅允许内部评估和测试，不能视为可分发的生产
依赖。项目不会自动安装全局依赖，也不会修改全局 Python/Node 环境。

精确版本见 `requirements-lock.txt`、`package.json` 和每次输出中的
`dependency_manifest.txt`。

## 8. 渲染、失败恢复和部分重跑

渲染优先级：

1. 命令行指定；
2. 环境变量；
3. PowerPoint COM 自动检测；
4. LibreOffice 独立临时配置目录；
5. 明确报告不可渲染。

从失败阶段继续：

```powershell
python run_ppt_workflow.py `
  --brief tests/fixtures/synthetic_brief.yaml `
  --input-root tests/fixtures/synthetic_project `
  --run-id 20260728_120000 `
  --resume `
  --from-stage generate
```

恢复必须使用同一 `run-id`，并保留相应的 `staging/<project>/<run-id>/` 与
`output/<project>/<run-id>/`。工作流会复用早于指定阶段的运行产物；缺失任何
必需产物时停止。

只更新部分页面的推荐做法：

1. 修改源文件或 brief；
2. 保留旧正式输出；
3. 创建新的 run-id；
4. 重新生成 storyboard 并人工对比 `slide_manifest.csv`；
5. 若需要模板内精确的逐页编辑，使用单独的修订任务，不覆盖原 PPTX。

## 9. QA 与人工审核

自动检查包括：

- PPTX ZIP 可打开；
- PowerPoint/LibreOffice 能真实渲染；
- PDF 页数等于 PPTX 页数；
- 每页 PNG 存在且非空白；
- 图层布局不超出画布（支持该后端时）；
- 输入哈希未变化；
- 每个正式 claim 具有有效 source_id；
- 必需引用不为空；
- 没有 `INFORMATION_REQUIRED` 被当作正式结论。

每次输出必须人工完成 `manual_review_checklist.md`。重点确认：

- 数字、单位、样本量、效应值、置信区间和 P 值；
- 观察性结果没有因果化；
- 计算推断没有升级为实验或临床证据；
- 阴性结果、局限性和不确定性没有遗漏；
- 引用真实、完整且能回溯；
- 没有患者级数据或直接标识符；
- 标题、图表和结论符合用户最终判断。

## 10. 状态解释

- `READY_FOR_ASSISTED_USE`：自动生成、渲染、溯源和最低 QA 均通过。
- `PARTIALLY_READY`：已生成部分工件，但存在可定位的科学、视觉或文件问题。
- `BLOCKED`：PPTX/PDF无法生成或打开，或关键输入/运行环境缺失。

三种状态都不代表最终科学批准。最终批准权始终属于用户。

## 11. 合成测试

运行单元测试：

```powershell
python -m unittest discover -s tests -v
```

运行完全合成的端到端测试：

```powershell
python run_ppt_workflow.py `
  --brief tests/fixtures/synthetic_brief.yaml `
  --input-root tests/fixtures/synthetic_project `
  --backend pptxgenjs
```

合成测试文件明确标注为虚构，不包含真实患者、机构、作者或外部文献。

## 12. 输出清单

每次正式输出至少包括：

- `<project>.pptx`
- `<project>.pdf`
- `preview/slide_001.png` 等
- `preview/contact_sheet.png`
- `deck_outline.md`
- `storyboard.csv`
- `speaker_notes.md`
- `source_manifest.csv`
- `slide_manifest.csv`
- `claim_source_map.csv`
- `figure_source_map.csv`
- `qa_report.md`
- `manual_review_checklist.md`
- `generation_log.md`
- `runtime_manifest.json`
- `dependency_manifest.txt`

所有正式输出位于新的时间戳目录，不覆盖既有结果。
