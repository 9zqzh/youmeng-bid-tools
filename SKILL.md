---
name: youmeng-bid-tools
description: 优盟投标工具：查询企业资质清单、把招标要求与企业资质逐条比对（满足/待补充/不满足/不适用）、对投标文件做废标项/无效标检查（逐片段滚动核查、风险分「已确认/待定」、导出 Excel 检查报告）、在写标书前设定「全局事实」并在写完后做全文一致性审计（跨节冲突 + 违反全局事实）、把流程图/思维导图/甘特图/架构图/柱状图等**配图渲染成 PNG**、用**导出模板**控制两册 Word 的排版（纸张/字体字号/标题样式/表格/页眉页脚/页码）、用**受限 HTML** 撰写有结构的正文（列表/子标题/表格/图注）、用**版面估算与字数配额**在出稿前发现篇幅问题（字数偏离、末页空洞）、生成投标文件商务技术册与报价册 Word 初稿，并可用「项目工作流」把 招标解读 → 全局事实 → 正文 → 一致性审计 → 两册 Word → 废标项检查 串成一条带关卡的可续跑链路。当用户要分析招标文件的资格/资质要求、核对企业资质能否满足投标条件、查找资质缺口、检查投标文件是否存在废标项/无效标风险、检查标书各章节数据前后矛盾、给标书画流程图或图表、调整标书 Word 排版或套用模板、控制标书篇幅与字数、完整走一遍投标文件编制流程、或依据招标要求起草投标文件 Word 时使用。
---

# 优盟标书工具

一个 **HTTP 瘦客户端**技能（仅用 Python 标准库），通过调用「优盟标书工具服务」（bid-mcp）完成：**列出资质 / 资质比对 / 废标项检查 / 全局事实 / 一致性审计 / 生成两册 Word**。

## 推荐用法：项目工作流（先看这一节）

如果你要**完整走一遍**「招标解读 → 全局事实 → 正文 → 一致性审计 → 两册 Word → 废标项检查」，别自己记那一堆 id —— 用项目工作流把它们收进一个 `project_id`：

```bash
python scripts/bid_tools.py project-init --project "XX项目" --bid-no "B-2026-001"
python scripts/bid_tools.py project-next --project-id <prj-id>       # 问「下一步做什么」
python scripts/bid_tools.py project-advance --project-id <prj-id> \
  --stage facts --artifacts '{"facts_id":"facts-20261015-xxxxxxxx"}'
python scripts/bid_tools.py project-state --project-id <prj-id>      # 看阶段状态与关卡结论
```

**关卡是硬校验的**（不是记个标记就算过），不达标会 `ok=false` 并告诉你具体缺什么：

| 阶段 | 必须带来的产物 | 放行条件 |
| --- | --- | --- |
| `analysis` | — | 招标解读与资质比对（起始阶段）|
| `facts` | `facts_id` | 事实清单非空、无未落实的占位符 |
| `content` | — | 前置阶段达标 |
| `consistency` | `ledger_id` | 台账已 `cons-repair` 给出统一口径、无未解决冲突 |
| `draft` | `files` | Word 文件确实存在于服务端 |
| `rejection` | `rejection_session_id` | 废标项检查已跑完（`finished`）|

确需跳过本阶段的拦阻项，可加 `--force`：会在项目里**留痕**（`history` 记录 `action: force` 与被跳过的原因），且阶段状态显示为 `accepted` 而不是 `done`；一旦该阶段产物内容再变化，豁免**自动失效**。
注意 `--force` 只豁免**当前阶段**的问题，前置阶段没达标仍会被拦住（要一级一级确认，避免一次把整条链路的门全踹开）。

**每个子命令也都能单独用**（下面的 1)–6) 节），项目工作流只是把它们串起来并加关卡，不影响单独调用。

## 前置

- **服务地址**：环境变量 `BID_API_BASE`，默认 `http://bid-mcp:18999`；跨机访问改为 `http://<服务器IP>:18999`。
- 仅需 **Python 3**（用标准库 `urllib`），**无需 pip 安装任何依赖**。
- 无需本机代码或资质文件（清单与原件由服务端持有）。

## 用法

脚本：`scripts/bid_tools.py`（输出 JSON）。

### 1) 列出企业资质

```bash
python scripts/bid_tools.py list
```

返回 `{company, items:[{id,name,category,valid_to,keywords}]}`。

### 2) 招标要求 × 资质 逐条比对

```bash
python scripts/bid_tools.py match --requirements '[{"id":"R1","text":"须具备质量管理体系认证（ISO 9001）","category":"资格","mandatory":true}]'
# 也可 --requirements @reqs.json
```

返回 `{company, results:[{requirement_id,status,matched_qualification_names,reason}], summary:{total,satisfied,unverified,missing,mandatory_missing,not_applicable}}`；
`status ∈ 满足 / 待补充 / 不满足 / 不适用`。

### 3) 生成两册 Word 初稿

```bash
python scripts/bid_tools.py draft \
  --project "江苏银行深圳分行宣传品供应商入围项目" \
  --sections '[{"title":"投标函","content":"致：（招标人）\n我方参加本项目投标。","volume":"business_technical"}]' \
  --qualifications q015,q006
```

返回 `{company, files:[{name,url}], warnings}`；`files[].url` 即下载链接（`kind` 由服务端按分册生成）。

### 4) 废标项 / 无效标检查（4 步，服务端维护权威状态）

**先由你（模型）从招标文件抽取「无效投标 / 废标项」清单**，再交给服务端建档：

```bash
python scripts/bid_tools.py reject-init \
  --project "XX项目" --bid-no "B-2026-001" \
  --items '[{"title":"未按第六章格式提交投标函","requirement":"投标函须按第六章格式填写"},{"title":"投标有效期不足90天","requirement":"投标有效期不少于90天"}]' \
  --documents '[{"name":"投标文件.pdf","chunk_total":6}]'
```

返回 `{session_id, checklist:[{id,type,title,scope,applicable}], documents, stats}`。
注意：纸质/线下事项（密封、纸质、现场递交、签章等）由服务端自动标为 `applicable=false`，不纳入电子检查。

**然后按「文档 × 片段」顺序逐段提交增量**（每次只交本片段的增量，**不要重发整份状态**）：

```bash
python scripts/bid_tools.py reject-apply --session rej-20261015-ab12cd34 --patch @patch.json
```

`patch` 的字段（全部可选，按需给）：

| 字段 | 含义 |
| --- | --- |
| `document_id` / `chunk_index` / `finished` | 游标：正在检查哪份文件的第几片段、是否全部完成 |
| `evidence_adds` | `[{checklist_id, locator, quote}]` 摘录到的投标原文证据 |
| `pending_risk_adds` | 疑似风险（含「没找到 / 未响应」类结论） |
| `pending_risk_updates` | `[{id, ...}]` 给待定项补证据；补齐后**自动升级**为已确认 |
| `pending_risk_resolves` | `[{id, resolution}]` 后续片段已排除该疑点 |
| `confirmed_risk_adds` | 直接确认的风险（必须带 `bid_evidence`） |

风险条目字段：`{checklist_id?, type?, severity: high|medium|low, title, requirement（检查依据）, bid_evidence（投标原文证据）, risk_reason, suggestion}`。

返回 `{round, applied:{...计数}, rejected:[{kind,reason}], stats, cursor}`。**务必检查 `rejected`**，按 reason 修正后重交。

```bash
python scripts/bid_tools.py reject-finalize --session <id>   # 收尾：去重排序，输出 findings/pending/summary
python scripts/bid_tools.py reject-export   --session <id>   # 导出 xlsx（检查概览/废标项/待定风险），返回下载链接
python scripts/bid_tools.py reject-state    --session <id>   # 查看当前完整状态（排查用）
```

## 三条硬规则（服务端强制执行，别试图绕过）

1. **「缺失 / 未响应」类结论不能直接确认**。服务端会把它降级为「待定」，因为文件后续片段可能包含对应内容。补到可引用的投标原文后，用 `pending_risk_updates` 升级。
2. **没有投标证据（`bid_evidence`）的条目不能确认**，只会被降级或拒绝。
3. **结构性线索保护**：若证据里出现「目录 / 章节标题 / 附件 / 清单 / 页码」等线索，不得据此判定缺失——同样降级为待定。

因此：**先广撒网进 `pending_risk_adds`，找到确凿原文再升级**，比一开始就硬报 `confirmed` 更有效。

### 5) 全局事实（**写正文之前**先做）

各节若各自生成，人员数、工期、质保期、金额、型号会互相打架。先固化一份权威事实清单：

```bash
python scripts/bid_tools.py facts-init \
  --project "XX项目" \
  --facts '[{"category":"日期与期限","name":"质保期","value":"一年"},{"category":"人员与数量","name":"项目经理","value":"1 人"}]'
```

返回 `{facts_id, facts, issues, rendered, ok}`：
- 服务端会**自动注入资质清单的企业主体信息**（名称 / 统一社会信用代码 / 法定代表人）作为**锁定事实**，你不得改写；提交冲突值会返回 `locked_conflict`。
- `issues` 里的 `placeholder`（含「待补充」等未落定值）与 `duplicate_conflict` **必须先处理**。
- `ok=false` 时不要把 `rendered` 拿去写正文。

人工评审后修正并回填：

```bash
python scripts/bid_tools.py facts-update --facts-id <id> --facts @facts.json
python scripts/bid_tools.py facts-render --facts-id <id>     # 取逐字节一致的注入块
```

**生成每一节正文时，把 `rendered` 完整注入**，并声明「与全局事实冲突即为错误」。

### 6) 一致性审计（**写正文之后**做）

你**并发核对每个小节**，只上报两类问题与本节的**类型化事实取值**；服务端负责比对：

```bash
python scripts/bid_tools.py cons-build --facts-id <id> --sections @sections.json --extractions @extractions.json
```

`extractions` 每项：`{section_id, facts:[{category,name,value,quote}], issues:[{type:"与全局事实冲突"|"小节内部矛盾", detail, quote}]}`。

返回 `{ledger_id, fact_table, conflicts, issues, summary}`。服务端会自动找出：
- **跨节同名不同取值**（如两节分别写「质保期一年 / 两年」）；
- **与全局事实不一致的取值**（锁定事实的违反会被标 `locked=true` 并优先列出）。

> **不要**上报用词、称谓、文风、篇幅、承诺语气强弱的差异——这些不进入比较，报了也只会被忽略。

再由你给出**统一口径**，服务端派发逐节修复任务：

```bash
python scripts/bid_tools.py cons-repair --ledger <id> --rules '[{"target":"质保期","value":"一年"}]'
```

返回 `{tasks:[{section_id,title,instructions}], unresolved, unused_rules, summary}`。
**`unresolved` 非空说明有冲突没给出口径**，必须补齐 `rules` 重跑，直到 `summary.ok=true`。
然后按 `tasks[].instructions` 逐节改写对应小节，改完可再跑一次 `cons-build` 复核。

### 7) 配图渲染（Mermaid / HTML → PNG）

流程图、思维导图、甘特图、架构图、柱状图这类图，交给服务端**确定性渲染**成 PNG。
（本期没有生图模型，实物照片 / 场景效果图 / 剖视图这类画不了，见 `fig-types`。）

```bash
python scripts/bid_tools.py fig-types              # 先看支持哪些类型、分别走哪种渲染方式
python scripts/bid_tools.py fig-status             # 环境自检：playwright / 浏览器 / mermaid.min.js
python scripts/bid_tools.py fig-render --figures @figs.json
```

`figs.json` 示例：

```json
[
  {"id": "flow", "type": "flowchart", "title": "评标流程",
   "mermaid": "flowchart TD\n  A[招标文件解析] --> B{资格性审查}\n  B -->|通过| C[技术方案编制]\n  B -->|不通过| D[废标]"},
  {"id": "bar", "type": "bar", "title": "人员配置",
   "html": "<div style=\"font-family:sans-serif\"><h3 style=\"color:#1a4d8f\">人员配置</h3><div style=\"display:flex;gap:8px;align-items:flex-end\"><div style=\"width:70px;height:120px;background:#4a90d9\"></div><div style=\"width:70px;height:70px;background:#7fb069\"></div></div></div>"}
]
```

返回 `{rendered:[{id,name,url,width,height}], failed:[{id,error}], summary}`。

**四条规则：**

1. `type` 可省，但给了就能自动决定渲染方式（`flowchart`/`流程图` → mermaid；`gantt`/`甘特图`/`bar`/`柱状图` → html）。
2. **html 只允许静态 HTML+CSS+SVG**：含 `<script>` / `<iframe>` / 内联事件（`onclick=`）/ `javascript:` 会被拒绝。图表请用**内联样式或手写 SVG** 画。
3. **渲染时网络已切断**：不能引用外部图片/字体/CSS（引了也不会加载），所以内容必须自包含。
4. **逐张容错**：某张失败不影响其它，`failed` 里带原因，改好后可只重交失败的那几张。

> 每张图的 `url` 可直接放进 Word 或给前端预览；图片落在服务端 `figures/` 子目录。

### 8) 导出模板（Word 排版）

两册 Word 的排版（纸张/页边距/字体字号/标题样式/表格/页眉页脚/页码格式）可以用模板控制：

```bash
python scripts/bid_tools.py template-presets                 # 看有哪些预设、各自什么排版
python scripts/bid_tools.py template-get --template a4-visual  # 解析成完整配置，便于微调
python scripts/bid_tools.py template-save --name house-style --template @tpl.json
python scripts/bid_tools.py draft --project "XX项目" --sections @s.json --template house-style
```

内置预设：

| 预设 | 排版 |
| --- | --- |
| `standard` | **默认**，与历史输出一致：A4 纵向、宋体 10.5、宋体黑色标题、仅页码 |
| `a4-visual` | A4 图文版：11pt / 1.5 倍行距 / 首行缩进、蓝色标题、页眉色条、页脚「第N页」 |
| `a3-landscape` | A3 横向：适合大表格与大架构图，图片可用宽度更大 |

**自定义模板**：`template-get` 拿到完整配置 → 改字段 → `template-save` 存成自己的名字 →
之后 `draft --template <名字>` 复用。可写 `base` 继承某个预设后再覆盖。

模板字段：`page`（paper/orientation/margins_cm/first_page_different）、`body`（font/size_pt/
line_spacing/first_line_indent_chars/align）、`heading`（font/color/size_pts[4]/bold/align/
space_before_pt/space_after_pt）、`table`（style/header_shading/font_size_pt/caption_prefix）、
`image`（max_width_cm/max_height_cm/align/caption_prefix）、`header` / `footer`
（enabled/text/align/size_pt/style/page_number_format）。

**字段取 `null` = 不修改**（沿用 Word 默认）。这正是 `standard` 能与历史输出逐字节一致的原因；
微调时别把不打算改的字段写成具体值。

**章节配图**：`sections[].images` 里放 `{"file": "flow.png", "title": "评标流程"}`，
`file` 用 `fig-render` 产出的文件名（也可给绝对路径），`title` 会自动编号成「图 N 标题」；
表格带 `title` 时同理编号为「表 N 标题」。**没给 title 就不会有题注**（避免给历史产物凭空加题注）。

### 9) 受限 HTML 正文（表达力最强，推荐）

`content` 是纯文本按换行拆段落，表达不了列表 / 子标题 / 表格 / 图注。要写**有结构**的正文，用 `sections[].html`：

```bash
python scripts/bid_tools.py html-spec                        # 先看白名单与示例
python scripts/bid_tools.py html-validate --html @body.html  # 出稿前自查
```

```json
[
  {"title": "技术方案", "volume": "business_technical",
   "html": "<p>本节说明<strong>实施要点</strong>。</p><h2>组织架构</h2><ul><li>项目经理 1 人</li></ul><table data-caption=\"人员配置\"><thead><tr><th>岗位</th><th>人数</th></tr></thead><tbody><tr><td>项目经理</td><td>1</td></tr></tbody></table><figure><img src=\"flow.png\"/><figcaption>评标流程</figcaption></figure>"},
  {"title": "报价说明", "volume": "price", "content": "见附表"}
]
```

| 支持 | 说明 |
| --- | --- |
| `p` / `br` | 段落与换行 |
| `h2` `h3` `h4` | 正文子标题（**`h1` 会被拒**：章节标题由 `sections[].title` 提供）|
| `ul` `ol` `li` | 项目符号 / 编号列表 |
| `table` + `thead`/`tbody`/`tr`/`th`/`td` | 表格；`table[data-caption]` 给表题，表头行用 `th`，`colspan` 支持 |
| `figure` + `img` + `figcaption` | 图片与图注；`img[src]` 用 `fig-render` 产出的图片名 |
| `strong` `b` `em` `i` `u` `s` `sup` `sub` | 行内格式 |

**明确不支持**（越界会**精确报错**，不会静默删掉）：`h1`、`rowspan`、`ol[start]`、`style` / `class`、
`script`/`iframe`/`svg` 等。题注只在给了 `data-caption` / `figcaption` 时生成并自动编号。

> `html` 与 `content` **二选一，html 优先**（同时给会按 html 渲染并在 warnings 里提示）。
> **务必先 `html-validate` 再出稿**：校验不过 `draft` 会直接报错（带具体是哪处、该怎么改），白跑一轮。

### 10) 版面与字数配额（出稿前自查）

两册 Word 的篇幅可以先用**确定性估算**把明显问题挡在出稿前，不必等出完稿翻 Word 才发现。

```bash
python scripts/bid_tools.py layout-metrics --template a4-visual   # 一页能装多少字
python scripts/bid_tools.py layout-plan --sections @sections.json --total-words 60000 \
  --min-words 800 --template a4-visual
```

`layout-plan` 返回三块：

| 块 | 内容 |
| --- | --- |
| `budget` | 按各节 `weight` 把 `--total-words` 分配成每节目标字数（`--min-words`/`--max-words` 约束）；某节自带 `target_words` 或给 `--section-words` 则直接采用 |
| `sections[]` | 每节字数、目标、偏离度、**估算页数**、**末页填充率**、图片/表格数、`flags` |
| `suggestions` | 可执行动作：扩写多少字 / 精简多少字 / 末页只填了 N% 建议调整字数或把图表挪到相邻小节 |

`flags` 取值：`字数不足` / `字数超出` / `页尾空洞` / `配图缺失`（阈值：偏离 20%、末页填充 < 35%）。
`sections` 与 `draft` 同构（`title` / `content` / `html` / `images` / `table`，可另加 `weight` / `target_words`）。

> **这是估算，不是 Word 实际分页。** 它由模板的纸张/页边距/字号/行距推出每页字数，图片高度按
> **真实像素**与模板画框换算。用途是出稿前发现明显问题，**不能替代在 Word 里核对版式**。
> 想让某节正好占满整页：把目标字数定在「页数 × 每页字数」附近即可。

## 典型工作流（项目驱动）
0. `project-init` 建项目；之后**每步先 `project-next` 问下一步、做完用 `project-advance` 交产物**；
1. `list` 了解企业有哪些资质 → 从招标文件提炼要求 → `match` 逐条比对，汇总「门槛缺口 / 待补充」→ advance 到 `analysis` 完成态；
2. **`facts-init` 固化全局事实**（处理掉 issues）→ advance 到 `facts`（带 `facts_id`）；
3. **`facts-render` 逐节注入** → 撰写各节正文（推荐用 `sections[].html` 写有结构的正文，先 `html-spec` 看白名单、`html-validate` 自查）；出稿前用 `layout-plan` 做**字数配额与版面自查**（字数偏离、末页空洞）；需要配图时用 `fig-render` 出图并在 html 里用 `<figure><img src="图名"/><figcaption>图题</figcaption></figure>` 引用 → advance 到 `content`；
4. **`cons-build` 建账 → `cons-repair` 给统一口径 → 按 `tasks` 修正各节** → advance 到 `consistency`（带 `ledger_id`）；
5. `draft` 生成 Word（需要特定排版就带 `--template <模板名>`，章节里可带 `images` 嵌入第 3 步渲染好的配图），把 `files[].url` 交给用户 → advance 到 `draft`（带 `files`）；
6. 从招标文件抽取废标项 → `reject-init` → 逐片段 `reject-apply` → `reject-finalize` → `reject-export`，把风险清单（尤其**已确认**项）与 xlsx 链接交给用户 → advance 到 `rejection`（带 `rejection_session_id`）。

若只想做其中一环（例如只查废标项），跳过 `project-*`，直接用对应的单个子命令即可。

## 注意

- **只对资质类要求比对**：流程 / 承诺 / 报价类要求会被标为「不适用」，不计入缺口。
- `draft` 必须提供 `--sections`；`--final` 且存在待补项时会拒绝导出。
- 报「无法连接 …」：检查 `BID_API_BASE` 是否指向可达的 bid-mcp 服务（沙箱网络需能访问该地址）。
- Word 的目录与页码需在 Word 中「更新域」；缺失的原件需人工补齐。
- 废标项检查的**待定项不是结论**，须人工核对原文；报告中的「已确认风险」才是有证据支撑的结论。
