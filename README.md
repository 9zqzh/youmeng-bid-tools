# 优盟标书工具 · youmeng-bid-tools

一个 **HTTP 瘦客户端技能**：查询企业资质清单、把招标要求与企业资质逐条比对、对投标文件做废标项/无效标检查、生成投标文件两册 Word 初稿。
**不依赖任何本机代码或资质文件**——只用 Python 标准库，通过 HTTP 调用「优盟标书工具服务」（bid-mcp）。

## 功能

| 子命令 | 说明 |
| --- | --- |
| `list`  | 列出企业资质清单（id / 名称 / 类别 / 有效期 / 关键词）|
| `match` | 招标要求 × 资质 **逐条比对**：满足 / 待补充 / 不满足 / 不适用 |
| `draft` | 生成 **商务技术册 + 报价册** Word 初稿（服务端嵌入资质原件）|
| `reject-init` | 创建**废标项检查会话**（传入从招标文件抽取的无效投标 / 废标项清单）|
| `reject-apply` | 提交**一个片段的检查增量**（证据 / 新增待定风险 / 更新 / 解除 / 确认）|
| `reject-finalize` | 收尾：合并去重、按严重度排序，输出报告 |
| `reject-export` | 导出 **xlsx 检查报告**（检查概览 / 废标项 / 待定风险）|
| `reject-state` | 查看会话当前完整状态（排查用）|
| `facts-init` | 建立**全局事实清单**（自动注入资质清单的企业主体信息作为锁定事实）|
| `facts-update` | 人工评审后按 id 全量替换事实清单 |
| `facts-render` | 取事实**注入块**（逐节生成正文时完整注入，保证各节一致）|
| `cons-build` | 一致性审计：把各节上报的类型化事实归集成台账，找出**跨节冲突**与**违反全局事实**的取值 |
| `cons-repair` | 一致性审计：按统一口径生成**逐节修复任务** |
| `project-init` | 创建**标书项目**，把全链路产物收拢到一个 `project_id` |
| `project-state` | 查看项目：当前阶段 / 各阶段状态（done / accepted / blocked / pending）/ 关卡结论 |
| `project-next` | 问「下一步该做什么」，直接给阶段指引与要带的产物 |
| `project-advance` | 声明某阶段完成并推进，**带关卡校验**（不达标拒绝；`--force` 可跳过并留痕）|
| `fig-types` | 列出可渲染的配图类型与渲染方式（画图前先看）|
| `fig-status` | 配图渲染环境自检（playwright / 浏览器 / mermaid.min.js）|
| `fig-render` | 把 **Mermaid / 静态 HTML** 渲染成 PNG，返回下载链接 |
| `template-presets` | 列出导出模板预设及其排版差异 |
| `template-list` | 列出已保存的自定义模板 |
| `template-get` | 把预设/模板名解析成完整配置（便于微调）|
| `template-save` | 保存自定义导出模板（可 `base` 继承预设）|
| `template-delete` | 删除自定义模板 |
| `html-spec` | 受限 HTML 正文的白名单与示例（写正文前先看）|
| `html-validate` | 校验一段受限 HTML（出稿前自查）|

## 目录结构

```
youmeng-bid-tools/
├── SKILL.md               # Skill 定义（名称、触发词、用法、硬规则）
└── scripts/
    └── bid_tools.py       # 全部子命令（纯标准库）
```

## 前置

- **Python 3**（仅标准库 `urllib`，**无需 pip 安装任何依赖**）。
- **服务地址**：环境变量 `BID_API_BASE`，默认 `http://bid-mcp:18999`；跨机访问改为 `http://<服务器IP>:18999`。
- 无需本机代码 / 资质文件（资质清单与原件由服务端持有）。

## 用法

```bash
# 1) 列出企业资质
python scripts/bid_tools.py list

# 2) 招标要求 × 资质 逐条比对（--requirements 支持内联 JSON 或 @文件）
python scripts/bid_tools.py match \
  --requirements '[{"id":"R1","text":"须具备质量管理体系认证（ISO 9001）","category":"资格","mandatory":true}]'

# 3) 生成两册 Word 初稿
python scripts/bid_tools.py draft \
  --project "江苏银行深圳分行宣传品供应商入围项目" \
  --sections '[{"title":"投标函","content":"致：（招标人）\n我方参加本项目投标。","volume":"business_technical"}]' \
  --qualifications q015,q006

# 4) 废标项检查：建档 → 逐片段提交 → 收尾 → 导出
python scripts/bid_tools.py reject-init \
  --project "XX项目" --bid-no "B-2026-001" \
  --items '[{"title":"未按第六章格式提交投标函","requirement":"投标函须按第六章格式填写"}]' \
  --documents '[{"name":"投标文件.pdf","chunk_total":6}]'

python scripts/bid_tools.py reject-apply --session rej-20261015-ab12cd34 --patch @patch.json

python scripts/bid_tools.py reject-finalize --session rej-20261015-ab12cd34
python scripts/bid_tools.py reject-export   --session rej-20261015-ab12cd34

# 5) 全局事实：写正文前固化，写完后做一致性审计
python scripts/bid_tools.py facts-init --project "XX项目" \
  --facts '[{"category":"日期与期限","name":"质保期","value":"一年"}]'

python scripts/bid_tools.py facts-render --facts-id facts-20261015-xxxxxxxx

python scripts/bid_tools.py cons-build --facts-id facts-20261015-xxxxxxxx \
  --sections @sections.json --extractions @extractions.json

python scripts/bid_tools.py cons-repair --ledger led-20261015-xxxxxxxx \
  --rules '[{"target":"质保期","value":"一年"}]'

# 6) 项目工作流：把全链路串起来（推荐入口）
python scripts/bid_tools.py project-init --project "XX项目" --bid-no "B-2026-001"
python scripts/bid_tools.py project-next  --project-id prj-20261015-xxxxxxxx
python scripts/bid_tools.py project-advance --project-id prj-20261015-xxxxxxxx \
  --stage facts --artifacts '{"facts_id":"facts-20261015-xxxxxxxx"}'
python scripts/bid_tools.py project-state --project-id prj-20261015-xxxxxxxx

# 7) 配图：先看支持哪些类型，再渲染
python scripts/bid_tools.py fig-types
python scripts/bid_tools.py fig-render --figures @figs.json

# 8) 导出模板：看预设 → 微调保存 → 出 Word 时套用
python scripts/bid_tools.py template-presets
python scripts/bid_tools.py template-get --template a4-visual
python scripts/bid_tools.py template-save --name house-style --template @tpl.json
python scripts/bid_tools.py draft --project "XX项目" --sections @s.json --template house-style

# 9) 受限 HTML 正文：先看白名单，再自查
python scripts/bid_tools.py html-spec
python scripts/bid_tools.py html-validate --html @body.html
```

输出均为 JSON：
- `match` → `{company, results:[{requirement_id,status,matched_qualification_names,reason}], summary:{...}}`
- `draft` → `{company, files:[{name,url}], warnings}`（`url` 即下载链接）
- `reject-init` → `{session_id, checklist, documents, stats}`
- `reject-apply` → `{round, applied:{...}, rejected:[{kind,reason}], stats, cursor}`
- `reject-finalize` → `{summary, findings, pending, resolved, checklist, warnings}`
- `reject-export` → `{name, url}`
- `facts-init` / `facts-update` → `{facts_id, facts, issues, rendered, ok}`
- `facts-render` → `{facts_id, rendered, count}`
- `cons-build` → `{ledger_id, fact_table, conflicts, issues, summary}`
- `cons-repair` → `{tasks:[{section_id,title,instructions}], unresolved, unused_rules, summary}`
- `project-init` / `project-state` / `project-advance` → `{project_id, stage, stages[], artifacts, blockers, warnings, can_advance, next}`
- `project-next` → `{action, stage, label, guide, advance_with}`
- `fig-types` → `{supported:[{type,label,kind}], unsupported:[...], kinds, hint}`
- `fig-status` → `{playwright, browser_configured, mermaid_js, ready, notes}`
- `fig-render` → `{rendered:[{id,name,url,width,height,bytes}], failed:[{id,error}], summary}`
- `template-presets` → `{presets:[{name,label,description,paper,orientation,body_size_pt}], fields}`
- `template-get` / `template-save` → `{name, label, template:{...}}`
- `template-list` → `{templates:[...], presets:[...]}`
- `html-spec` → `{block_tags, heading_tags, inline_tags, attributes, not_supported, example, rules}`
- `html-validate` → `{ok, issues:[{level,kind,message,tag,attr}], stats, message}`

## 受限 HTML 正文

`sections[].content` 是纯文本按换行拆段落，表达不了列表 / 子标题 / 表格 / 图注。
要写有结构的正文就用 `sections[].html`，由程序**校验白名单后确定性转 Word**（不经 LibreOffice 往返）。

| 支持 | 说明 |
| --- | --- |
| `p` / `br` | 段落与换行 |
| `h2` `h3` `h4` | 正文子标题（`h1` 会被拒：章节标题由 `sections[].title` 提供）|
| `ul` `ol` `li` | 项目符号 / 编号列表 |
| `table` + `thead`/`tbody`/`tr`/`th`/`td` | 表格；`table[data-caption]` 给表题，`colspan` 支持 |
| `figure` + `img` + `figcaption` | 图片与图注；`img[src]` 用 `fig-render` 产出的图片名 |
| `strong` `b` `em` `i` `u` `s` `sup` `sub` | 行内格式 |

**越界精确报错、不静默删掉**：`h1` / `rowspan` / `ol[start]` / `style` / `class` / `script` 等都会返回
带位置与修改建议的 issue。所以**先 `html-validate` 再出稿**；校验不过 `draft` 会直接报错。

`html` 与 `content` **二选一，html 优先**（同时给会按 html 渲染并在 warnings 提示）。

## 导出模板（Word 排版）

两册 Word 的排版可以用模板控制，不必改代码：

| 预设 | 排版 |
| --- | --- |
| `standard` | **默认**，与历史输出**逐字节一致**：A4 纵向、宋体 10.5、宋体黑色标题、仅页码 |
| `a4-visual` | A4 图文版：11pt / 1.5 倍行距 / 首行缩进、蓝色标题、页眉色条、页脚「第N页」 |
| `a3-landscape` | A3 横向：大表格与大架构图友好，图片可用宽度更大 |

**自定义**：`template-get` 拿完整配置 → 改字段 → `template-save --name house-style` →
`draft --template house-style` 复用。可写 `base` 继承预设。

**字段取 `null` = 不修改**（沿用 Word 默认）—— 这就是 `standard` 能与历史输出逐字节一致的原因。
支持的字段：`page` / `body` / `heading` / `table` / `image` / `header` / `footer`（详见 `template-presets` 的 `fields`）。

**章节配图与题注**：`sections[].images = [{"file":"flow.png","title":"评标流程"}]`，
`file` 用 `fig-render` 产出的文件名；有 `title` 才生成题注，自动编号「图 N 标题」/「表 N 标题」。
页码是**可刷新的 PAGE 域**（不是写死的数字），Word 里按 F9 更新。

## 配图渲染（Mermaid / HTML → PNG）

标书里的图绝大多数是**结构图与图表**，用 Mermaid 与 HTML+SVG **确定性渲染**比生图更可控、可复现、也更清晰。

| 渲染方式 | 适用类型 |
| --- | --- |
| `mermaid` | 思维导图 / 流程图 / 实体关系图 / 时序图 / 状态图 / 类图 |
| `html` | 甘特图 / 组织架构 / 泳道 / RACI / 风险矩阵 / 系统架构 / 拓扑 / WBS / 鱼骨 / 柱状 / 折线 / 饼 / 雷达 / 面积 / 散点 / 时间轴 |
| 不支持 | 实物照片 / 场景效果图 / 抠图 / 剖视图 / 爆炸分解图 / 三维渲染（**需生图模型**，本期不做）|

**四条硬规则**（服务端强制）：

1. `type` 可省，但给了就自动决定渲染方式（也接受中文标签，如「甘特图」）。
2. **html 只允许静态 HTML+CSS+SVG**：含 `<script>` / `<iframe>` / 内联事件 / `javascript:` 会被拒绝。
3. **渲染时网络已切断**：不能引用外部图片、字体、CSS（引了也不会加载），内容必须自包含。
4. **逐张容错**：单张失败不影响其它，`failed` 里带原因，可只重交失败的。

图片落在服务端 `figures/` 子目录，通过 `/files/figures/<name>.png` 下载；`rendered[].url` 可直接使用。
渲染环境自检用 `fig-status`：它会告诉你 playwright 是否可用、浏览器来源、`mermaid.min.js` 是否就位。

## 项目工作流（把全链路串成一条带关卡的可续跑链路）

问题：整条链路每个阶段的产物是一个 id（`facts_id` / `ledger_id` / `files` / `rejection_session_id`），
全靠模型自己记、自己保证顺序。一旦漏做「全局事实」就直接写正文，各节数据必然互相打架，而且要到成品才暴露。

做法：把这些产物收进一个 `project_id`，并在 `project-advance` 时**真的去校验**产物是否达标。

| 阶段 | 产物 | 放行条件 |
| --- | --- | --- |
| `analysis` | — | 起始阶段 |
| `facts` | `facts_id` | 事实清单非空、无未落实占位符 |
| `content` | — | 前置阶段达标 |
| `consistency` | `ledger_id` | 已给统一口径、无未解决冲突 |
| `draft` | `files` | Word 文件确实存在于服务端 |
| `rejection` | `rejection_session_id` | 废标项检查已 `finished` |

阶段状态：`done`（达标）/ `accepted`（曾 `--force` 跳过，产物仍未达标，显式标出）/ `blocked`（未达标）/ `pending`。
`--force` 只豁免**当前阶段**；且豁免绑定产物**内容指纹**，内容一变即失效；前置阶段未达标要一级一级确认。

## 全局事实与一致性审计的设计取舍

**为什么把比对放到服务端**：几十节的标书不可能塞进一次模型调用去「通读比较」。所以让模型**并发逐节抽取类型化事实**（人员与数量 / 日期与期限 / 金额 / 技术参数…），程序把各节事实按 (分类, 名称) 归集成台账，再判冲突。既可靠、可复现，也能覆盖长文档。

**去噪是设计出来的，不是靠提示词**：只比较类型化事实，用词、称谓、文风、篇幅、承诺语气强弱**根本不进入比较**，因此不会产生噪音问题。

**事实分类**：企业主体 / 人员与数量 / 日期与期限 / 频次与时限 / 地点与范围 / 金额 / 技术参数 / 编号与名称 / 责任归属 / 其他。

**锁定事实**：`facts-init` 会自动把资质清单里的企业名称 / 统一社会信用代码 / 法定代表人注入为**锁定事实**，模型提交冲突值会被标 `locked_conflict` 并以清单为准。

**空值与占位符**：含「待补充 / 待确认 / XXX」等未落定值会被 `lint` 标出，不能靠它写正文。

## 废标项检查的硬规则

服务端（不是本脚本）维护权威状态，并对模型输出做强制约束：

1. **「缺失 / 未响应」类结论不能直接确认** → 自动降级为「待定」（后续片段可能推翻）。
2. **没有 `bid_evidence` 的条目不能确认**。
3. **结构性线索保护**：证据含「目录 / 章节标题 / 附件 / 清单 / 页码」时不得判定缺失。
4. 纸质/线下事项（密封、纸质、现场递交、签章等）自动标为 `applicable=false`，排除出检查范围。

所以正确姿势是：**先广撒网进 `pending_risk_adds`，拿到确凿原文再用 `pending_risk_updates` 升级**。

## 作为 Skill 安装

- **项目级**：放到 `<repo>/.qoder/skills/youmeng-bid-tools/`
- **用户级**：放到 `~/.qoder-cn/skills/youmeng-bid-tools/`
- **WeKnora**（沙箱技能）：把本仓库打成 zip 上传到「设置 → 技能沙箱」，并在**智能推理**智能体的「能力扩展」里启用 Skills 并勾选本技能。

## 服务端接口（本技能调用）

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET  | `/api/qualifications` | 资质清单 |
| POST | `/api/qualify-match`   | 要求 × 资质比对 |
| POST | `/api/generate-draft`  | 生成两册 Word，返回下载链接 |
| POST | `/api/rejection/init`  | 创建废标项检查会话 |
| POST | `/api/rejection/apply` | 提交一个片段的检查增量 |
| POST | `/api/rejection/finalize` | 收尾并输出报告 |
| POST | `/api/rejection/export` | 导出 xlsx 报告 |
| GET  | `/api/rejection/state` | 查看会话状态 |
| POST | `/api/facts/init` | 建立全局事实清单（注入锁定事实）|
| POST | `/api/facts/update` | 全量替换事实清单 |
| POST | `/api/facts/render` | 取事实注入块 |
| POST | `/api/consistency/build` | 建立一致性台账、找冲突 |
| POST | `/api/consistency/repair` | 生成逐节修复任务 |
| POST | `/api/project/init` | 创建标书项目 |
| GET  | `/api/project/state` | 项目状态（阶段 / 产物 / 关卡）|
| GET  | `/api/project/next` | 下一步该做什么 |
| POST | `/api/project/advance` | 声明阶段完成并推进（关卡校验）|
| GET  | `/api/project/stages` | 阶段定义 |
| GET  | `/api/figure/types` | 可渲染的配图类型 |
| GET  | `/api/figure/status` | 渲染环境自检 |
| POST | `/api/figure/render` | 渲染 Mermaid / HTML 为 PNG |
| GET  | `/api/template/presets` | 导出模板预设 |
| GET  | `/api/template/list` | 用户模板列表 |
| POST | `/api/template/get` | 解析预设/模板为完整配置 |
| POST | `/api/template/save` | 保存自定义模板 |
| POST | `/api/template/delete` | 删除自定义模板 |
| GET  | `/api/html/spec` | 受限 HTML 白名单与示例 |
| POST | `/api/html/validate` | 校验受限 HTML |
| GET  | `/files/{name}`        | 下载生成的 docx / xlsx / png（png 在 figures/ 子目录）|

## 说明与边界

- **只对资质类要求比对**：流程 / 承诺 / 报价类要求会被标为「不适用」，不计入缺口。
- `draft` 必须提供 `--sections`；`--final` 且存在待补项时会拒绝导出。
- 报「无法连接 …」：检查 `BID_API_BASE` 是否指向可达的 bid-mcp 服务。
- Word 的目录与页码需在 Word 中「更新域」；缺失的原件需人工补齐。
- 废标项检查的**待定项不是结论**，须人工核对原文；只有「已确认风险」带证据支撑。
