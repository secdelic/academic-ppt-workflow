# Academic PPT Workflow：快速使用

本工作流把日常操作压缩为四条路线和三个质量等级。用户只需准备项目资料、填写内容brief，并在需要时批准视觉brief；科学证据、页面规划、渲染与QA由仓库统一执行。

## 1. 只需选择路线和质量

路线：

- `generate`：从资料生成新Deck。
- `enhance`：在已有PPT上增加或替换1–10页；默认继承原Deck风格。
- `template-fill`：填充正式机构PPTX/POTX；Master、Layout和品牌规则优先。
- `template-create`：从获授权的风格参考创建可复用模板配置。

质量：

- `quick`：快速草稿，仅保留必要输出与受影响页面QA。
- `validated`：日常正式工作推荐，增加PDF、contact sheet及最小科学/视觉审计。
- `full`：冻结版本、首次高风险项目或最终完整验证。

## 2. 新建项目

在仓库根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/new_project.ps1 `
  -ProjectId "my-study" `
  -Route generate `
  -Quality validated
```

脚本会在 `<PPT_WORKSPACE_HOME>/projects/my-study/` 创建标准目录，并输出项目完整路径。它不会覆盖已有项目文件。

## 3. 放入资料并填写brief

- 文档：`input/documents/`
- 表格：`input/data/`
- 科学图片：`input/figures/`
- 参考文献：`input/references/`
- 获授权风格参考：`input/style_reference/`
- 正式模板：`input/template/`
- 待更新PPT：`input/existing/`

填写 `brief/presentation_brief.yaml`，重点确认受众、时长、页数、`must_include`、`must_not_claim`、隐私模式及模板/既有PPT路径。

## 4. 什么时候需要GPT视觉总监

- 新建Deck或重大视觉重设计：使用 [GPT_Art_Director.md](../prompts/GPT_Art_Director.md) 生成3套方向。选择后，把获批方案保存为 `brief/visual_brief.yaml`。
- 已有PPT小改：默认跳过GPT。不要创建视觉brief，系统继承原Deck。
- 正式模板填充：模板品牌规则不可被视觉brief覆盖；GPT只建议模板允许的自由区域。

初始化生成的 `visual_brief.DRAFT.yaml` 只是草案，不会生效。人工批准并另存为 `visual_brief.yaml` 后才生效。

## 5. 运行

```powershell
python run_ppt_workflow.py --project "<项目完整路径>"
```

如需临时覆盖brief中的选择：

```powershell
python run_ppt_workflow.py --project "<项目完整路径>" `
  --route enhance --quality quick
```

## 6. 找到结果

每次运行使用独立run目录：

```text
output/<run_id>/draft/
output/<run_id>/final/
output/<run_id>/preview/
audit/<run_id>/
```

`quick`以Draft和受影响页面QA为主；`validated/full`至少提供PPTX、PDF、contact sheet及最小科学/视觉审计。自动通过不等于最终科学批准，交付前仍需人工科学与视觉审核。

`template-create`是配置型路线：它在`draft/style_profile.json`交付获授权参考的只读风格配置，不伪造PPTX、PDF或contact sheet。需要实际Deck时，请随后使用`template-fill`或`generate`。
