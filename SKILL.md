---
name: youmeng-bid-tools
description: 优盟投标工具：查询企业资质清单、把招标要求与企业资质逐条比对（满足/待补充/不满足/不适用）、生成投标文件商务技术册与报价册 Word 初稿。当用户要分析招标文件的资格/资质要求、核对企业资质能否满足投标条件、查找资质缺口、或依据招标要求起草投标文件 Word 时使用。
---

# 优盟标书工具

一个 **HTTP 瘦客户端**技能（仅用 Python 标准库），通过调用「优盟标书工具服务」（bid-mcp）完成：**列出资质 / 资质比对 / 生成两册 Word**。

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

## 典型工作流

1. `list` 了解企业有哪些资质；
2. 从招标文件（或已解析文本）提炼要求 → `match` 逐条比对，汇总「门槛缺口 / 待补充」；
3. 依据要求与命中资质撰写简短 `sections` → `draft` 生成 Word，把 `files[].url` 交给用户。

## 注意

- **只对资质类要求比对**：流程 / 承诺 / 报价类要求会被标为「不适用」，不计入缺口。
- `draft` 必须提供 `--sections`；`--final` 且存在待补项时会拒绝导出。
- 报「无法连接 …」：检查 `BID_API_BASE` 是否指向可达的 bid-mcp 服务（沙箱网络需能访问该地址）。
- Word 的目录与页码需在 Word 中「更新域」；缺失的原件需人工补齐。
