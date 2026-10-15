# 优盟标书工具 · youmeng-bid-tools

一个 **自包含技能**：查询企业资质清单、把招标要求与企业资质逐条比对、对投标文件做废标项/无效标检查、检查错别字与前后矛盾、给标书配图、控制排版与篇幅、多份标书查重、生成投标文件两册 Word 初稿。
**不依赖任何外部服务**——实现随技能一起分发（`scripts/bidagent/`），在技能沙箱里就地跑。
**唯一的权威源在后端仓库 `services/`**，技能里的副本由 `tools/sync_from_backend.py` 同步生成，两边不会漂移。

## 功能

| 子命令 | 说明 |
| --- | --- |
| `list`  | 列出企业资质清单（id / 名称 / 类别 / 有效期 / 关键词）|
| `match` | 招标要求 × 资质 **逐条比对**：满足 / 待补充 / 不满足 / 不适用 |
| `draft` | 生成 **商务技术册 + 报价册** Word 初稿（嵌入资质原件）|
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
| `fig-render` | 把 **Mermaid / 静态 HTML** 渲染成 PNG，落到产物目录 |
| `template-presets` | 列出导出模板预设及其排版差异 |
| `template-list` | 列出已保存的自定义模板 |
| `template-get` | 把预设/模板名解析成完整配置（便于微调）|
| `template-save` | 保存自定义导出模板（可 `base` 继承预设）|
| `template-delete` | 删除自定义模板 |
| `html-spec` | 受限 HTML 正文的白名单与示例（写正文前先看）|
| `html-validate` | 校验一段受限 HTML（出稿前自查）|
| `layout-metrics` | 算出一套模板下每页能装多少字 |
| `layout-plan` | 字数配额 + 版面估算 + 偏差与建议（出稿前自查篇幅）|
| `dup-analyze` | **多份投标文件查重**（招标白名单剔除 + 句级/目录/元数据/图片四层比对 + 候选对）|
| `typo-check` | 回验错别字（**定位不到的直接排除**，招标既有用词降级为待核对）|
| `logic-check` | 逻辑谬误（跨段前后矛盾）|
| `check-export` | 把 **废标项/错别字/逻辑** 导出到**同一份 xlsx** |

## 目录结构

```
youmeng-bid-tools/
├── SKILL.md                     # Skill 定义（名称、触发词、用法、硬规则）
├── requirements.txt             # Python 依赖（装到技能自己的 .venv）
├── scripts/
│   ├── bid_tools.py             # 全部 33 个子命令（命令行入口，输出 JSON）
│   ├── assets/
│   │   └── mermaid.min.js       # 离线渲染 Mermaid 用（已随技能带）
│   └── bidagent/                # 实现本体：从后端 services/ 同步而来（勿手改）
│       ├── settings.py          #   ← 技能侧手写的配置垫片（不参与同步）
│       ├── bid_api.py           #   业务入口层（CLI 直接调用它）
│       ├── bid_service.py …     #   各功能模块
│       └── .synced.json         #   后端 commit + 各文件 sha256（供 --check 溯源）
└── tools/
    ├── sync_from_backend.py     # 把后端 services/ 同步进 scripts/bidagent/
    └── smoke_test.py            # 装好后自检：依赖 + 各子命令
```

## 前置（装一次）

**不需要任何服务**，但要先把依赖装进技能自己的虚拟环境：

```bash
uv venv --seed .venv
uv pip install --python .venv/bin/python -r requirements.txt
python tools/smoke_test.py          # 自检：依赖 + 各子命令能否跑通
```

- 依赖：`python-docx / openpyxl / lxml / Pillow / pypdf / pypdfium2 / httpx`；
- **配图渲染**额外需要 `playwright` + Chromium/Chrome/Edge + 中文字体（否则中文变方块），见 `fig-status`。
- 本地已有 Python 时也可 `python -m venv .venv && .venv/bin/pip install -r requirements.txt`。
- **不要 `pip install` 进系统 Python**（沙箱里会污染环境，也会在技能升级时丢失）。

三个目录（有默认值，一般不用改；用同名环境变量覆盖）：

| 环境变量 | 用途 | 默认 |
| --- | --- | --- |
| `BID_INPUT_DIR` | **输入**：`manifest.json` 与资质原件放这里 | `/workspace/input` |
| `BID_OUTPUT_DIR` | **产物**：Word / xlsx / png 落这里 | `/workspace/output` |
| `BID_STATE_DIR` | **状态**：检查会话 / 事实 / 台账 / 项目 / 模板 | `<技能目录>/.state` |
| `BID_MANIFEST_PATH` | 资质清单路径（留空则在输入目录里找 `manifest.json`） | 空 |

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

# 10) 版面与字数配额：出稿前自查篇幅
python scripts/bid_tools.py layout-metrics --template a4-visual
python scripts/bid_tools.py layout-plan --sections @sections.json --total-words 60000 --min-words 800

# 11) 查重：多份投标文件比对（务必带上招标原文，否则会大面积误报）
python scripts/bid_tools.py dup-analyze --documents @docs.json --tender @tender.txt

# 12) 错别字与逻辑谬误（三类检查可合并导出同一份 xlsx）
python scripts/bid_tools.py typo-check --candidates @typos.json --source @bid.txt --tender @tender.txt
python scripts/bid_tools.py logic-check --segments @segments.json --extractions @extractions.json
python scripts/bid_tools.py check-export --rejection-session <id> --typo-id <id> --logic-id <id>
```

输出均为 JSON：
- `match` → `{company, results:[{requirement_id,status,matched_qualification_names,reason}], summary:{...}}`
- `draft` → `{company, files:[{name,path}], warnings, output_dir}`（`path` 为落盘绝对路径）
- `reject-init` → `{session_id, checklist, documents, stats}`
- `reject-apply` → `{round, applied:{...}, rejected:[{kind,reason}], stats, cursor}`
- `reject-finalize` → `{summary, findings, pending, resolved, checklist, warnings}`
- `reject-export` → `{name, path, output_dir}`
- `facts-init` / `facts-update` → `{facts_id, facts, issues, rendered, ok}`
- `facts-render` → `{facts_id, rendered, count}`
- `cons-build` → `{ledger_id, fact_table, conflicts, issues, summary}`
- `cons-repair` → `{tasks:[{section_id,title,instructions}], unresolved, unused_rules, summary}`
- `project-init` / `project-state` / `project-advance` → `{project_id, stage, stages[], artifacts, blockers, warnings, can_advance, next}`
- `project-next` → `{action, stage, label, guide, advance_with}`
- `fig-types` → `{supported:[{type,label,kind}], unsupported:[...], kinds, hint}`
- `fig-status` → `{playwright, browser_configured, mermaid_js, ready, notes}`
- `fig-render` → `{rendered:[{id,name,path,width,height,bytes}], failed:[{id,error}], summary, output_dir}`
- `template-presets` → `{presets:[{name,label,description,paper,orientation,body_size_pt}], fields}`
- `template-get` / `template-save` → `{name, label, template:{...}}`
- `template-list` → `{templates:[...], presets:[...]}`
- `html-spec` → `{block_tags, heading_tags, inline_tags, attributes, not_supported, example, rules}`
- `html-validate` → `{ok, issues:[{level,kind,message,tag,attr}], stats, message}`
- `layout-metrics` → `{template, label, metrics:{chars_per_line, lines_per_page, chars_per_page, usable_cm, note}}`
- `layout-plan` → `{metrics, budget:{mode,target_total,sections}, sections:[...flags], summary, suggestions, note}`
- `dup-analyze` → `{signature, summary, content, outline:{pairs}, metadata:{documents,findings}, images, candidates, tender_whitelist, warnings, note}`
- `typo-check` → `{review_id, findings:[{status,location,line_text,wrong,right,...}], unlocated, invalid, summary, warnings, note}`
- `logic-check` → `{review_id, findings:[{type,category,name,segments,distinct,detail}], register, summary, warnings, note}`
- `check-export` → `{name, path, sheets, output_dir}`

## 错别字与逻辑谬误检查

两件事，都遵循同一原则：**模型提线索，程序回验**。

**错别字**：你给出候选，工具回原文查证 —— 模型对专业术语很容易"造错"，所以定位不到的**直接排除**。

| 状态 | 含义 |
| --- | --- |
| `confirmed` | 原文已定位，且该文本不在招标原文里 → 可以改 |
| `needs_review` | 该文本**出现在招标原文里**，可能是招标方既有用词 → 人工确认 |
| `unlocated` | **原文里找不到，疑似误判，已排除** |

同一个错别字出现多次时，会优先用你给的 `context` 锁定**它自己那一次**（定位点始终指向错别字本身）。
正文与招标都支持给文件路径（`--source-path` / `--tender-path`），大文档走路径更稳。

**逻辑谬误**：逐段上报事实 → 程序按 (分类, 名称) 归集登记 → **跨段取值不一致即判前后矛盾**。
只比类型化事实（人数/工期/金额/技术参数…），用词与文风差异不进入比较。

**合并导出**：`check-export` 把三类结果写进**同一份 xlsx** —— `检查概览 / 废标项 / 待定风险 / 错别字 / 逻辑问题`。

## 标书查重（确定性引擎）

比对多份投标文件的雷同，**不调用模型**，结果可复核可复现。

```bash
python scripts/bid_tools.py dup-analyze --documents @docs.json --tender @tender.txt
```

**`--tender`（招标原文）强烈建议提供。** 招标须知 / 格式要求 / 评标办法会被各家大段照抄，
不做剔除会**大面积误报**成"雷同"；命中招标来源的句子会单独归类为「招标原文引用」，不计入统计。

| 层 | 内容 |
| --- | --- |
| `content` | 句子级雷同（归一化精确 + 二元组 Dice 近重复）+ 招标白名单剔除量 |
| `outline` | 目录相似度（路径 0.45 / 标题 0.35 / 顺序 0.2）与 high/medium/low 风险档 |
| `metadata` | 同源取证：整文件 sha256、docx 的 rsid 编辑会话指纹、PDF trailer ID / 增量更新 / 签名数 |
| `images` | 图片字节 sha256 复用 |

`documents` 每项 `{id?, name?, text?|path?}`：给 `text` 只做文本层；给 `path` 才做元数据与图片层。
`candidates` 形如 `{current_source_id, reference_source_id, similarity, overlap_excerpt}`，
可直接喂给解释型模型节点。`signature` 是输入指纹 —— **内容一变就该把旧结论标记为失效**。

> **边界**：相似度**不等于法律侵权或围串标证据**，只说明文字/结构层面重合；
> 「法定固定格式 vs 实质重复」的定性判断交给模型，引擎只给可复核的数值与原文摘录。

## 版面与字数配额

出稿前用**确定性估算**把明显问题挡住，不必等翻 Word 才发现篇幅不对。

```bash
python scripts/bid_tools.py layout-metrics --template a4-visual
python scripts/bid_tools.py layout-plan --sections @sections.json --total-words 60000 --min-words 800
```

| 块 | 内容 |
| --- | --- |
| `budget` | 按各节 `weight` 把总字数分配成每节目标（受 `min_words`/`max_words` 约束）；某节自带 `target_words` 或给 `section_words` 则直接采用 |
| `sections[]` | 每节字数 / 目标 / 偏离度 / **估算页数** / **末页填充率** / 图片表格数 / `flags` |
| `suggestions` | 扩写多少字、精简多少字、末页只填了 N% 建议调整字数或把图表挪到相邻小节 |

`flags`：`字数不足` / `字数超出` / `页尾空洞` / `配图缺失`（阈值：偏离 20%、末页填充 < 35%）。

> **是估算，不是 Word 实际分页。** 由模板的纸张/页边距/字号/行距推算每页字数，图片高度按
> **真实像素**与模板画框换算。用于出稿前自查，**不能替代在 Word 里核对版式**。
> 想让某节正好占满整页：目标字数定在「页数 × 每页字数」附近。

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

**四条硬规则**（程序强制）：

1. `type` 可省，但给了就自动决定渲染方式（也接受中文标签，如「甘特图」）。
2. **html 只允许静态 HTML+CSS+SVG**：含 `<script>` / `<iframe>` / 内联事件 / `javascript:` 会被拒绝。
3. **渲染时网络已切断**：不能引用外部图片、字体、CSS（引了也不会加载），内容必须自包含。
4. **逐张容错**：单张失败不影响其它，`failed` 里带原因，可只重交失败的。

图片落在**产物目录**的 `figures/` 子目录（`rendered[].path`）；章节 `images` 里按文件名引用。
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
| `draft` | `files` | Word 文件确实已落盘到产物目录 |
| `rejection` | `rejection_session_id` | 废标项检查已 `finished` |

阶段状态：`done`（达标）/ `accepted`（曾 `--force` 跳过，产物仍未达标，显式标出）/ `blocked`（未达标）/ `pending`。
`--force` 只豁免**当前阶段**；且豁免绑定产物**内容指纹**，内容一变即失效；前置阶段未达标要一级一级确认。

## 全局事实与一致性审计的设计取舍

**为什么把比对交给程序**：几十节的标书不可能塞进一次模型调用去「通读比较」。所以让模型**并发逐节抽取类型化事实**（人员与数量 / 日期与期限 / 金额 / 技术参数…），程序把各节事实按 (分类, 名称) 归集成台账，再判冲突。既可靠、可复现，也能覆盖长文档。

**去噪是设计出来的，不是靠提示词**：只比较类型化事实，用词、称谓、文风、篇幅、承诺语气强弱**根本不进入比较**，因此不会产生噪音问题。

**事实分类**：企业主体 / 人员与数量 / 日期与期限 / 频次与时限 / 地点与范围 / 金额 / 技术参数 / 编号与名称 / 责任归属 / 其他。

**锁定事实**：`facts-init` 会自动把资质清单里的企业名称 / 统一社会信用代码 / 法定代表人注入为**锁定事实**，模型提交冲突值会被标 `locked_conflict` 并以清单为准。

**空值与占位符**：含「待补充 / 待确认 / XXX」等未落定值会被 `lint` 标出，不能靠它写正文。

## 废标项检查的硬规则

**程序（不是模型）维护权威状态**，并对模型输出做强制约束：

1. **「缺失 / 未响应」类结论不能直接确认** → 自动降级为「待定」（后续片段可能推翻）。
2. **没有 `bid_evidence` 的条目不能确认**。
3. **结构性线索保护**：证据含「目录 / 章节标题 / 附件 / 清单 / 页码」时不得判定缺失。
4. 纸质/线下事项（密封、纸质、现场递交、签章等）自动标为 `applicable=false`，排除出检查范围。

所以正确姿势是：**先广撒网进 `pending_risk_adds`，拿到确凿原文再用 `pending_risk_updates` 升级**。

## 作为 Skill 安装

- **项目级**：放到 `<repo>/.qoder/skills/youmeng-bid-tools/`
- **用户级**：放到 `~/.qoder-cn/skills/youmeng-bid-tools/`
- **WeKnora**（沙箱技能）：把本仓库打成 zip 上传到「设置 → 技能沙箱」，并在**智能推理**智能体的「能力扩展」里启用 Skills 并勾选本技能。
  上传后在沙箱里装一次依赖（会话内即可）：

  ```bash
  cd "$WEKNORA_SKILL_DIR" && uv venv --seed .venv && \
    uv pip install --python .venv/bin/python -r requirements.txt && \
    python tools/smoke_test.py
  ```

  没有 `uv` 时退回 `python -m venv .venv && .venv/bin/pip install -r requirements.txt`。
  配图渲染还要 `playwright install chromium` + 中文字体（`fonts-noto-cjk`）。

## 实现与同步（这套实现从哪来）

本技能的实现**不是在这里手写的**，而是与后端服务**共用同一份代码**：

```
后端仓库 digital-hub-test/services/*.py      ← 唯一权威源（服务端 bid-mcp 也用它）
            │  tools/sync_from_backend.py（改写 import + 记账 sha256）
            ▼
技能 scripts/bidagent/*.py                   ← 生成物，勿手改
```

- 想改逻辑：**改后端 `services/`**，再跑 `python tools/sync_from_backend.py` 重新生成；
  用 `python tools/sync_from_backend.py --check` 可以查两边是否已漂移。
- 只有 `scripts/bidagent/settings.py` 是技能侧手写的配置垫片（纯标准库读同名环境变量，不参与同步）。
- `.synced.json` 记录了同步时的后端 commit 与各文件 sha256，供溯源与 `--check` 使用。

这样做的收益：**服务端与沙箱里的 skill 行为逐字节一致**，改一处就两边都对；不会出现"服务器修了 bug，技能里还是老版本"。
`bid-mcp` 服务**仍然保留**（给走 HTTP 的调用方与 WeKnora 的 MCP 工具用），只是**不再是本技能的依赖**。

| 消费方式 | 入口 | 传输 |
| --- | --- | --- |
| 技能沙箱（本技能） | `scripts/bid_tools.py` | 无（直接 import `bidagent.bid_api`）|
| MCP 工具 | `mcp-server/mcp_server.py` | FastMCP（streamable-http）|
| REST | `/api/*` | HTTP，返回下载链接 |

三者共用 `services/bid_api.py` 的业务入口层，所以参数、返回体字段、硬规则完全一致；
差别只在"产物怎么给"：服务端返回**下载链接**，技能本地返回**落盘路径**。

## 说明与边界

- **只对资质类要求比对**：流程 / 承诺 / 报价类要求会被标为「不适用」，不计入缺口。
- `draft` 必须提供 `--sections`；`--final` 且存在待补项时会拒绝导出。
- 报「找不到资质清单」：把 `manifest.json` 放进 `BID_INPUT_DIR`（默认 `/workspace/input`）。
- 报「缺 xxx 模块 / 命令找不到」：依赖没装好，先跑 `python tools/smoke_test.py`。
- 状态持久化在 `BID_STATE_DIR`（默认技能目录 `.state/`）；清理状态即等于重来。
- Word 的目录与页码需在 Word 中「更新域」；缺失的原件需人工补齐。
- 废标项检查的**待定项不是结论**，须人工核对原文；只有「已确认风险」带证据支撑。
