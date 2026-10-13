# 优盟标书工具 · youmeng-bid-tools

把标书业务能力封装为一个 **Agent Skill**：查询企业资质清单、把招标要求与企业资质逐条比对、生成投标文件两册 Word 初稿。
与同名 MCP 服务功能一致——**复用后端仓库 `digital-hub` 的 `services` 逻辑**（非重复实现）。

## 功能

| 子命令 | 说明 |
| --- | --- |
| `list`  | 列出企业资质清单（id / 名称 / 类别 / 有效期 / 关键词）|
| `match` | 招标要求 × 资质 **逐条比对**：满足 / 待补充 / 不满足 / 不适用 |
| `draft` | 生成 **商务技术册 + 报价册** Word 初稿（可嵌入资质原件）|

## 目录结构

```
youmeng-bid-tools/
├── SKILL.md               # Skill 定义（名称、触发词、用法）
├── scripts/
│   └── bid_tools.py       # list / match / draft 三个子命令
└── assets/
    ├── manifest.json      # 企业资质结构化清单
    └── 优识企业资质/       # 15 份资质原件（供 Word 嵌入）
```

## 依赖

- Python 3.10+
- 后端仓库 `digital-hub` 的依赖：`pip install -r requirements.txt`
  （python-docx / pypdf / pypdfium2 / Pillow / httpx / pydantic-settings）

> ⚠️ 脚本会**自动定位后端仓库根**，因此本仓库需位于后端仓库内：
> `<digital-hub>/.qoder/skills/youmeng-bid-tools/`。
> 若单独克隆本仓库使用，请用环境变量 **`BID_REPO`** 指向 `digital-hub` 仓库根。

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
  --qualifications q015,q006 \
  --out ./out
```

输出均为 JSON：
- `match` → `{company, results:[{requirement_id,status,matched_qualification_names,reason}], summary:{...}}`
- `draft` → `{company, files:[{name,path,kind}], warnings}`（`kind`：`business_technical` / `price`）

## 作为 Skill 安装

- **项目级**（随后端仓库）：放到 `<digital-hub>/.qoder/skills/youmeng-bid-tools/`。
- **用户级**（本机全局）：放到 `~/.qoder-cn/skills/youmeng-bid-tools/`，并设置 `BID_REPO` 指向后端仓库。

## 说明与边界

- **只对资质类要求比对**：流程 / 承诺 / 报价类要求会被标为「不适用」，不计入缺口。
- `draft` 必须提供 `--sections`；`--final` 且存在待补项时会拒绝导出。
- Word 的目录与页码需在 Word 中「更新域」；缺失的原件（如营业执照、财务/业绩）需人工补齐。
- 更新资质清单或原件后，请同步 `assets/` 目录。
