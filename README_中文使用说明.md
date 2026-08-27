# Academic PPT Workflow 中文使用说明

本仓库是本地优先、来源可追溯、人工最终批准的学术 PowerPoint 生产系统。
仓库是唯一 canonical authority；生产渲染仍只使用 native PptxGenJS。

当前版本是私有发布候选，不代表已经获准公开发布，也不代表自动 QA 可以替代
科学审核和最终视觉审核。

## 日常只需要两个选择

- 路线：`generate`、`enhance`、`template-fill`、`template-create`
- 质量：`quick`、`validated`、`full`

用户不需要准备 SlideSpec、VisualSpec、operation plan、内容哈希或缓存检查点。
这些内部对象由工作流管理。

## Windows 首次安装

```powershell
$env:PPT_WORKFLOW_HOME = (Resolve-Path .).Path
$env:PPT_WORKSPACE_HOME = Join-Path $env:LOCALAPPDATA "AcademicPPTWorkspace"
powershell -ExecutionPolicy Bypass -File scripts/bootstrap_windows.ps1
powershell -ExecutionPolicy Bypass -File scripts/doctor.ps1
```

安装脚本只创建项目级 Python/Node 环境，不修改全局依赖。

## 新建并运行项目

```powershell
powershell -ExecutionPolicy Bypass -File scripts/new_project.ps1 `
  -ProjectId "my-study" `
  -Route generate `
  -Quality validated

.\.venv\Scripts\python.exe run_ppt_workflow.py --project "<项目目录>"
```

正式输入与输出必须位于配置的外部 workspace。不要把患者资料、真实临床 PPT、
private cache、运行日志或密钥放进仓库。

详细步骤见：

- [快速使用](docs/README_%E5%BF%AB%E9%80%9F%E4%BD%BF%E7%94%A8.md)
- [用户使用说明书](docs/%E7%94%A8%E6%88%B7%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E%E4%B9%A6.md)
- [首次安装](README_FIRST_INSTALL.md)

只有隐私审计、合成 CI、依赖锁和发布包验证全部通过，并由人工确认 GitHub
仓库为 Private 后，才允许继续建立 tag 或 Release。
