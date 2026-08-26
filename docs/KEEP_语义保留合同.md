# KEEP 语义保留合同

v2.7.0-rc4 将 KEEP 页面保留判定分为三个同时成立的 Gate：

1. `PRESENTATION_IDENTITY`：SlideID、最终顺序、Master/Layout 目标、媒体关系语义、notes 角色和 protected designation；
2. `SEMANTIC_EDITABLE_IDENTITY`：文本、格式、几何、层级、表格、图片、图表、动作及其他可编辑语义的 canonical manifest；
3. `RENDER_IDENTITY`：raw OOXML 不同但前两项一致时，使用同一 PowerPoint 环境逐页导出并进行无正文遮罩的精确像素比较。

`raw_shape_tree_hash` 继续记录，但仅为 `DIAGNOSTIC_ONLY`。未知差异默认失败；只有 `config/keep_normalization_allowlist.yaml` 中逐项登记且已有 before/after 证据的字段可以归一化。

RC3 的完全合成表格页在 PowerPoint SaveAs 后出现以下非语义变化：序列化空白和属性顺序、空 `effectLst`、表格 row/column bookkeeping、默认垂直边距省略及内部 shape id 规范化。扩展的多类型合成回归还确认：PowerPoint 会在媒体二进制和裁剪完全相同时重命名或去重内部 media part。上述行为均要求语义 manifest 一致且 render identity 通过。

下列真实变化不得归一化：正文或表格 cell 文字、字体、cell fill、border、行高/列宽、对象几何、z-order、媒体二进制、Master/Layout、chart data。相关负向 fixture 均必须失败。

RC3 历史状态保留为：本机与 GitHub CI 通过，第二设备部分通过；失败原因是未修改表格页经过 PowerPoint COM SaveAs 后 raw shape-tree hash 产生假阳性。RC3 不移动、不覆盖，由 RC4 候选修复取代。
