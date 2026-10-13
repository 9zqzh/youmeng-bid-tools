---
name: youmeng-bid-tools
description: 优盟投标工具：查询企业资质清单、把招标要求与企业资质逐条比对（满足/待补充/不满足/不适用）、生成投标文件商务技术册与报价册 Word 初稿。当用户要分析招标文件的资格/资质要求、核对企业资质能否满足投标条件、查找资质缺口、或依据招标要求起草投标文件 Word 时使用。
---

# 优盟标书工具

复用后端仓库（digital-hub）的 `services` 逻辑，提供三件事：**列出企业资质清单 / 资质逐条比对 / 生成两册 Word 初稿**。功能与同名 MCP 服务一致。

## 前置

- 本 skill 位于后端仓库内：`<repo>/.qoder/skills/youmeng-bid-tools/`；脚本自动定位仓库根。若仓库在别处，用环境变量 `BID_REPO` 指定。
- 依赖后端仓库 `requirements.txt`（python-docx / pypdf / pypdfium2 / Pillow / httpx / pydantic-settings）：
  `pip install -r <repo>/requirements.txt`
- 企业资质清单与原件已随 skill 打包在 `assets/`（`manifest.json` + `优识企业资质/`）。

## 用法

脚本：`scripts/bid_tools.py`（用仓库的 Python 运行，输出为 JSON）。

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

返回：

```json
{"company":"...","results":[{"requirement_id":"R1","status":"满足","matched_qualification_names":["质量管理体系认证证书（ISO 9001:2015）"],"reason":"..."}],
 "summary":{"total":1,"satisfied":1,"unverified":0,"missing":0,"mandatory_missing":0,"not_applicable":0}}
```

`status ∈ 满足 / 待补充 / 不满足 / 不适用`。

### 3) 生成两册 Word 初稿

```bash
python scripts/bid_tools.py draft \
  --project "江苏银行深圳分行宣传品供应商入围项目" \
  --sections '[{"title":"投标函","content":"致：（招标人）\n我方参加本项目投标。","volume":"business_technical"}]' \
  --qualifications q015,q006 \
  --out ./out
```

- `--sections`：每项 `{title,content,volume}`，`volume` 取 `business_technical`（商务技术册）或 `price`（报价册）；
- `--qualifications`：要嵌入扫描件的资质 id（见 `list`），逗号分隔；
- 返回 `{company, files:[{name,path,kind}], warnings}`；`kind ∈ business_technical / price`。

## 典型工作流

1. 运行 `list` 了解企业有哪些资质；
2. 从招标文件（或已解析文本）提炼要求 → 用 `match` 逐条比对，汇总「门槛缺口 / 待补充」；
3. 依据要求与命中资质撰写简短 `sections` → 用 `draft` 生成 Word，把 `files[].path` 交给用户。

## 注意

- **只对资质类要求比对**：流程 / 承诺 / 报价类要求会被标为「不适用」，不计入缺口。
- `draft` 必须提供 `--sections`（不提供则不生成正文）；`--final` 且存在待补项时会拒绝导出。
- Word 目录与页码需在 Word 中「更新域」；缺失的原件（如营业执照、财务/业绩）需人工补。
- 资质原件位于本 skill 的 `assets/优识企业资质/`；更新清单或原件后需同步该目录。
