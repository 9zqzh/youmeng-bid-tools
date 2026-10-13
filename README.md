# 优盟标书工具 · youmeng-bid-tools

一个 **HTTP 瘦客户端技能**：查询企业资质清单、把招标要求与企业资质逐条比对、生成投标文件两册 Word 初稿。
**不依赖任何本机代码或资质文件**——只用 Python 标准库，通过 HTTP 调用「优盟标书工具服务」（bid-mcp）。

## 功能

| 子命令 | 说明 |
| --- | --- |
| `list`  | 列出企业资质清单（id / 名称 / 类别 / 有效期 / 关键词）|
| `match` | 招标要求 × 资质 **逐条比对**：满足 / 待补充 / 不满足 / 不适用 |
| `draft` | 生成 **商务技术册 + 报价册** Word 初稿（服务端嵌入资质原件）|

## 目录结构

```
youmeng-bid-tools/
├── SKILL.md               # Skill 定义（名称、触发词、用法）
└── scripts/
    └── bid_tools.py       # list / match / draft 三个子命令（纯标准库）
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
```

输出均为 JSON：
- `match` → `{company, results:[{requirement_id,status,matched_qualification_names,reason}], summary:{...}}`
- `draft` → `{company, files:[{name,url}], warnings}`（`url` 即下载链接）

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
| GET  | `/files/{name}`        | 下载生成的 docx |

## 说明与边界

- **只对资质类要求比对**：流程 / 承诺 / 报价类要求会被标为「不适用」，不计入缺口。
- `draft` 必须提供 `--sections`；`--final` 且存在待补项时会拒绝导出。
- 报「无法连接 …」：检查 `BID_API_BASE` 是否指向可达的 bid-mcp 服务。
- Word 的目录与页码需在 Word 中「更新域」；缺失的原件需人工补齐。
