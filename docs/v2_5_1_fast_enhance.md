# Academic PPT Workflow v2.5.1：Fast Enhance 使用与治理说明

## 1. 目的与证据边界

`fast_enhance` 用于已有 PPTX 的小规模增量更新：通常修改、替换或新增
1–10 页。它不增加新的科学证据层，也不改变以下 v2.5 权威合同：

```text
SourceRegistry
  -> CanonicalEvidenceGraph
  -> StoryGraph
  -> ArtDirectionSpec
  -> SlideSpec
  -> VisualSpec
  -> Renderer
```

缓存只是上述对象的经验证快照，不是第二套科学权威。无法证明影响范围时，
必须升级到 `full_validation`。临床模式仍要求人工科学批准，自动 QA 不得替代
临床判断。

## 2. 两种运行模式

| 模式 | 适用范围 | 默认行为 |
|---|---|---|
| `full_validation` | 第一次真实项目、科学定义变化、大规模重生成、版本冻结或正式全面验证 | 保留原完整解析、渲染与 QA 链；所有非 `enhance-existing` 路线默认使用 |
| `fast_enhance` | 已验证项目且候选页/operation plan已准备的小规模增量更新 | `enhance-existing` 路由可选用；只解析改变文件、重建受影响对象、修改受影响页 |

以下任一情况不得使用 fast 路径：

- 没有已经通过完整验证并晋级的项目缓存；
- cohort、暴露/干预、结局、baseline、时间窗、统计模型、缺失数据策略或验证策略改变；
- source 被删除；
- 影响图缺少 `source -> claim -> visual -> slide` 绑定；
- master、全 Deck 语言、主题或 ArtDirection 发生无法局部证明的变化；
- operation plan 与源 PPTX、候选 PPTX 或 stable `slide_id` 不一致。

## 3. Project Cache 与 Delta

Production中，应用目录与用户工作区分离。缓存权威优先级为：显式内部
`project_cache_root`、`PPT_CACHE_HOME`、最后是项目目录内的默认值：

```text
<PPT_WORKSPACE_HOME>/projects/<PROJECT_ID>/cache/<opaque_project_key>/
```

如配置共享缓存根，则使用：

```text
<PPT_CACHE_HOME>/<opaque_project_key>/
```

仓库内`.cache/project_state/`仅用于`DEV_TEST_ONLY`和已有兼容测试，不是
Production默认值，也不能覆盖公共Project Interface解析出的项目缓存。

临床模式强制 `local_private_cache_only`。缓存：

- 只使用 UTF-8 JSON，不使用 pickle/joblib/marshal；
- 用 HMAC 派生 opaque project key，不以患者姓名、床号、文件名或项目标题命名；
- 以 SHA-256 作为 delta 的唯一改变权威；mtime 和 size 只可作为提示；
- generation 不可变；通过同一文件系统内的原子 pointer promote 切换 current；
- 保留 previous generation，失败或中断不得晋级半成品；
- 拒绝 UNC 以及 `tests/`、`regression/`、`benchmark/`、`output/`、
  `archive/` 中的缓存根。
- fast 复用前同时核对完整 brief、科学定义、extractor、config 和运行代码的
  SHA-256 指纹；任一缺失或变化都 fail closed 到 `full_validation`；
- fast 成功后从新 PPTX 重新扫描 master/layout、theme font 和 media registry；
  扫描失败时明确标记失效，不沿用旧 deck 的资产状态；
- 临床 `full_validation` 在创建 project-key secret 或 cache 目录之前复用同一
  canonical containment gate；brief 中的 clinical/MDT 声明同样触发。项目外
  只允许显式配置的`PPT_CACHE_HOME`，仓库内缓存还必须通过`.gitignore` gate。

Delta 只有四种状态：

- `UNCHANGED`：SHA-256 相同，不重新解析；
- `MODIFIED`：同一 source identity 的 SHA-256 改变；
- `NEW`：当前新增；
- `REMOVED`：当前删除，fast 路径升级为完整验证。

## 4. ChangeImpactGraph 与增量修改

影响图使用 stable ID 和显式边：

```text
source -> claim -> visual -> slide
```

它输出 `affected_slides`。缺边、未知 ID、removed entity 或 Deck-wide change
均返回 `FULL`，不能猜测局部范围。

`operation_plan.json` 是确定性 PowerPoint 操作合同，支持：

- `KEEP`
- `REPLACE`
- `INSERT_AFTER`

未改变页保持原 OOXML、master/layout 和媒体。PowerPoint 在工作副本上执行一次
增量操作，不能覆盖输入文件。当前 fast 路径要求显式提供原 PPTX、operation plan
及 changed-slide candidate PPTX；它不会自行重新生成整个 Deck，也尚未在 production
入口内自动完成临床/科研内容整合、候选页规划与生成。因此，只有这些上游工件均已
source-bound 审核时才能使用；否则必须 fail closed 到 `full_validation`。

## 5. QA 与交付 Gate

### Changed-slide QA

只对新增或修改页运行：

- scientific wording boundary；
- source binding；
- PowerPoint actual geometry、safe-zone、overflow；
- typography；
- visual QA。

### Whole-deck lightweight QA

每次 fast 运行仍检查：

- PowerPoint 可打开；
- 页数与 slide order；
- relationships 和 media；
- master/footer 一致；
- 明显空白页。

### Gate

- `DRAFT_READY`：changed-slide QA 通过、source binding 通过、几何通过、PPTX
  可打开；普通 fast 运行最高到此状态。
- `FINAL_DELIVERY_READY`：显式 `--final-delivery`，完成全 Deck PowerPoint QA，
  且用户显式提供 `--manual-review-approved`。不得从自动 QA 推断人工批准。

LibreOffice 默认关闭，仅 `--cross-renderer-validation` 时运行。`--audit-full`
仅控制完整审计副本，不放宽隐私规则。

## 6. 命令与 flags

首次真实项目或缓存重建：

```powershell
py -3.13 run_ppt_workflow.py `
  --route enhance-existing `
  --workflow-mode full_validation `
  --brief brief/presentation_brief.yaml `
  --existing-pptx <source.pptx> `
  --previous-deck-ir <previous_deck_ir.json> `
  --update-slide <stable_slide_id>
```

`full_validation` 的 enhance route 仍沿用 v2.x 合同：必须在
`--update-slide`、`--update-section`、`--update-figure` 中选择且只选择一个，
并提供 previous Deck IR（或在源 Deck 同目录提供 `deck_ir.json`）。

普通 fast draft：

```powershell
py -3.13 run_ppt_workflow.py `
  --route enhance-existing `
  --workflow-mode fast_enhance `
  --brief brief/presentation_brief.yaml `
  --existing-pptx <source.pptx> `
  --operation-plan <operation_plan.json> `
  --changed-candidate-pptx <changed_slides.pptx> `
  --clinical-privacy-mode
```

可选 flags：

| Flag | 行为 |
|---|---|
| `--export-pdf` | fast draft 额外导出一次完整 PDF |
| `--final-delivery` | 启用最终全 Deck PowerPoint gate |
| `--manual-review-approved` | 只与 `--final-delivery` 同时使用，记录用户人工批准 |
| `--cross-renderer-validation` | 显式启用 LibreOffice 页数比较 |
| `--audit-full` | 复制完整审计中间件；仍不得复制私有 cache |
| `--retry-slide <slide_id>` | 使用只含该页的隔离 operation plan 重试一页 |
| `--retry-stage <stage>` | 只重跑 allowlist 中的失败阶段 |
| `--resume --run-id <id>` | 在已有 staging/output run 内续跑 |

`--retry-slide` 与 `--retry-stage` 互斥。可重试阶段当前为
`changed_slide_generation`、`powerpoint_apply`、`changed_slide_render`、
`changed_slide_qa`、`whole_deck_light_qa`。

## 7. 默认最小产物

用户交付合同为：

1. `updated.pptx`
2. `change_plan.json`
3. `slide_diff.json`
4. `changed_slide_qa.md`
5. `manual_review_checklist.md`
6. `execution_summary.md`

changed-slide PNG、PowerPoint 几何数据、delta、影响图和 lightweight QA 默认放在
staging。PDF 只有显式请求或 final delivery 时生成。

运行时 profile 是治理证据，但不增加正式交付数量。最终合成 COM 验证确认默认
output 恰好包含上述 6 项；`runtime_profile.json` 保存在 staging/runtime evidence
边界内。

## 8. Runtime Budget 与基线

复杂度目标：

| 等级 | changed slides | 目标 |
|---|---:|---:|
| S | 0–3 | ≤15 分钟 |
| M | 4–10 | ≤30–45 分钟 |
| L | ≥11 | ≤60–90 分钟 |

匿名历史样本墙钟时间为 223.779 分钟。其中最大区间是 candidate remediation
loop：126.464 分钟，占 56.5%。9 次 PowerPoint 增量编辑累计仅 135.651 秒；
主要瓶颈是反复全 Deck 诊断、渲染、LibreOffice、几何检查、修复与包装，而不是
单次 COM 修改。

M 级 fast run 超过 60 分钟时必须生成 `runtime_bottleneck_report.md`，分别报告
文件解析、生成/规划、PowerPoint、QA 和模型耗时。没有可靠模型遥测时必须写
`NOT_MEASURED`，不能估算调用次数或 token。

## 9. 已验证与未验证

截至本文件生成：

- 全仓回归：241/241 通过；
- fast 定向回归：64/64 通过；
- routing/full compatibility：34/34 通过；
- cache delta、原子 generation、影响图 escalation、operation plan、changed-slide
  QA、lightweight QA、fast 路由、retry 控制、runtime profiler 与纯合成 fixture
  合同已运行验证；
- retry stage 已实际执行并验证只从目标阶段恢复；
- Windows PowerPoint COM 合成 M 级 apply/QA run：外层22.081577秒、内部
  21.692225秒，21→25 页，7 个
  changed/new（3 REPLACE + 4 INSERT）、18 KEEP；
- source parse：2；cache reuse：1；PowerPoint 调用：2；LibreOffice：0；
  model calls：0；token usage：0；
- KEEP semantic identity：18/18；PowerPoint 可打开；input hashes 不变；
- 默认 output：恰好 6 项；
- 真实病例内容未进入测试、benchmark、regression 或本文档；
- fast run 下真实科学正确性与人工 review：`NOT_ASSESSED`；
- routing/full compatibility 合同已验证，但生产材料上的完整 `full_validation`
  实跑仍为 `NOT_ASSESSED`。

该计时中的候选页和operation plan由确定性合成fixture提前准备；真实科学内容整合
与高推理模型不在计时边界内。因此不能把22秒与旧223.779分钟完整任务直接比较。
基于当前证据，工程状态为 `FAST_ENHANCE_OPTIMIZATION_PARTIAL`：增量apply/QA引擎
达到目标，但 production planner/candidate generation、修订候选后的单页retry及
真实端到端性能仍待闭合。临床内容与最终视觉仍需人工审核。
