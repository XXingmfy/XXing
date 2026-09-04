# 图与图源规范（结构图必须可编辑矢量源）

> 论文中除"系统运行截图"外的所有图（用例图、业务/时序流程图、总体架构图、功能模块结构图、E-R 图、技术路线图等结构图）**必须出自可编辑矢量图源**，并在 `figure-registry.yaml` 登记源文件。**禁止把脚本直接涂出的位图 PNG 当作论文结构图交付**（学校/导师会要求提供 Visio 原图或可编辑源）。

## 1. 图的两类与对应要求

| 图类 | 例子 | 交付要求 |
|---|---|---|
| 运行截图 | 小程序页面截图、后台界面截图 | 真实 PNG/JPG（最终完成态、无脏数据），**允许且只能是位图** |
| 结构图 | 用例图、流程图、时序图、架构图、模块结构图、E-R 图、技术路线图 | **必须可编辑矢量源**：首选 Visio `.vsdx`；可用 draw.io `.drawio`（可一键另存 `.vsdx`）；docx 内嵌其高清导出图（PNG/SVG 转 PNG），并把源文件随论文留档 |

判定：一张结构图如果只有 `.png` 且无任何可编辑源（`.vsdx`/`.drawio`/`.vssx`）登记，视为未达标，在交付报告中列为证据缺口。

## 2. 生成工具链（按环境选择，优先从上到下）

### 2.1 本机装有 Microsoft Visio（Windows）
用 Visio COM 自动化（PowerShell 或 python `win32com`）按数据生成 `.vsdx`：
1. 用结构化描述（实体/用例/流程步骤/模块树）驱动布局；
2. 打开 Visio → 按数据建形状与连线 → 另存 `.vsdx`；
3. 导出 PNG 用于 docx 插图；有 Word 同机 Visio 时可将 `.vsdx` 以 OLE 嵌入 docx（Word 中可双击编辑）。
（hbst-thesis 不内置整套 Visio COM 脚本，按需在目标机编写；布局与重叠断言可参考 scripts/gen_diagram.py 的思路。）

### 2.2 本机装有 draw.io（免费，推荐无 Visio 时使用）
1. 用 `scripts/gen_diagram.py` 从结构化描述生成可编辑源 `xxx.drawio`；
2. 用 draw.io 命令行导出论文插图 PNG 与可编辑 `.vsdx`：
   ```bash
   drawio -x -f png -o figs/fig3_1_usecase.png figs/fig3_1_usecase.drawio
   drawio -x -f vsdx -o figs/fig3_1_usecase.vsdx figs/fig3_1_usecase.drawio   # 供交审/留档
   ```
   （若本机 draw.io CLI 未入 PATH，用桌面版打开 `.drawio` → File → Export as → PNG / VSDX。）
3. `.drawio` 是公开 XML，**Visio 2013+ 可直接打开 `.vsdx`**，或 draw.io 内 File→Export as→VSDX 得到 Visio 原生文件。

### 2.3 两者皆无
- 仍用 `gen_diagram.py` 产出 `.drawio` 可编辑源；
- docx 内嵌图可用脚本自绘的 PNG **仅作临时预览**，`figure-registry.yaml` 中该图 `status: needs_visio_export`、`source_format: drawio`；
- 交付报告中明确列出"需在装有 draw.io / Visio 的环境导出 .vsdx 的图清单"。

## 3. 各结构图的规范要点

- **用例图（第 3 章）**：系统边界框 + 参与者(人形/方框) + 用例椭圆；用例名称与正文需求措辞**完全一致**。
- **业务/时序流程图**：开始/结束(圆角或椭圆)、处理(矩形)、判断(菱形)、箭头带方向；分支有明确条件文字；无交叉重叠（用分层或正交走线）。
- **总体架构图**：分层（端 → 后台 → 服务 → 数据库/第三方），每层列出模块/技术，箭头表示调用/数据流。
- **功能模块结构图**：树状分层，模块名与正文/后台菜单一致。
- **E-R 图**：总 E-R 只放实体+关系（菱形/连线标注 1:n）；字段进"单实体 E-R"与三线表。实体/关系名与表名、正文一致。
- 结构图文字：中文黑体/宋体统一、清晰不重叠；导出 PNG 建议 ≥200dpi、宽度适配版心（10–15cm）。

## 4. figure-registry.yaml 图条目字段（结构图必填源信息）

```yaml
- id: fig3_1_usecase
  caption: 图3.1　系统总体用例图
  chapter: 3
  kind: usecase            # usecase | flowchart | seq | arch | module | er | screenshot
  source_file: figs/fig3_1_usecase.drawio   # 可编辑源(.vsdx/.drawio)
  source_format: drawio    # vsdx | drawio
  export_png: figs/fig3_1_usecase.png
  status: ready            # ready | needs_user_screenshot | needs_visio_export
  first_mention: "第3.1节"
```

## 5. 与写作/降AIGC衔接
- 正文只写"如图 X.Y 所示"，不写图源格式；图源信息只存在于 registry 与交付报告。
- 结构图的"可编辑源缺失"属于证据缺口，须在交付报告列出，**不得用"AI 工作流画图"措辞遮掩**（04-thesis-voice）。
