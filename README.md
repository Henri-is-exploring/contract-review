# contract-review · 合同审核 Skill

![type](https://img.shields.io/badge/type-AI%20Skill-orange)
![python](https://img.shields.io/badge/python-3.9%2B-blue)
![format](https://img.shields.io/badge/format-.docx-green)

一个把「资深法务审合同」做成可复用工作流的 AI Skill。上传 `.docx` 合同后，自动走完 **立场确认 → 知识库审查 → 产出 Word 修订稿 + 审核报告 → 审核要点回补知识库** 的全流程。

> 设计原则：**修订准确 > 流程完整 > 速度**。宁可慢、宁可多问，也不许跳步、不许假改。

## 📦 直接装进你的 Agent

这是一个 Agent 技能（Skill），采用通用的 `SKILL.md` + 脚本结构，无需任何改造，放入你所使用 Agent 的 skills 目录即可：

- 把整个 `contract-review/` 文件夹复制到你的 Agent 的 skills 目录下（目录名保持 `contract-review`）；
- 具体路径取决于你用的 Agent，例如某些 Agent 的 skills 目录形如 `~/.agent/skills/` 或应用数据下的 `skills/` 文件夹；
- 若你的 Agent 支持从 Git 直接加载，也可 `git clone` 到对应目录。

放入后，在对话里上传一份 `.docx` 合同并说「审核这份合同」，Agent 会先确认你的立场（甲方 / 乙方）、谈判地位与审核尺度，再按下方 8 步工作流产出 Word 修订稿 + 审核报告。

## 特性

- **立场驱动**：先确认我方是甲方/乙方、谈判地位、审核尺度，再逐条审。
- **对抗视角**：假设对方严格按条款行权、我方处于最不利情形来挑风险。
- **真实修订**：产出带 `w:ins` / `w:del` / 批注的 Word 修订稿（真实 OOXML 元素，非样式模拟）。
- **自我学习**：把本次可复用的检查点回补进知识库（只追加、不改删）。

## 目录结构

```
contract-review/
├── SKILL.md                           # 工作流编排（8 步）
├── references/
│   └── contract-review-playbook.md    # 核心审查知识库（只追加不修改）
└── scripts/
    ├── extract_clauses.py             # 抽条款清单，防跳读
    ├── track_changes.py               # 生成 Word 修订稿（真实修订标记 + 批注）
    ├── generate_report.py             # 生成审核报告 .docx
    └── update_playbook.py             # 审核要点回补知识库（含去重）
```

## 工作流（8 步）

1. 接收合同 + 确认立场（甲方/乙方）
2. 总结业务图景 + 列待确认点
3. 向用户确认（谈判地位 / 审核尺度 / 业务图景 + 全部待确认点）
4. 基于知识库逐项审查（强制先抽条款清单防跳读）
5. 生成修订合同（真实 OOXML 修订标记）
6. 生成审核报告（整体风险等级 / 条款建议 / 待确认事项）
7. 回补知识库 + 未命中复盘
8. 交付三份产出

## 快速开始

```bash
# 1) 依赖
pip install python-docx

# 2) 抽取条款清单（防跳读，逐条审查前必跑）
python3 scripts/extract_clauses.py --input 合同.docx --output clauses.json

# 3) 按审查结果生成修订稿 / 报告 / 回补知识库（见下方示例）
python3 scripts/track_changes.py      --input 合同.docx --edits edits.json        --output 修订合同.docx
python3 scripts/generate_report.py    --input review_result.json                 --output 审核报告.docx
python3 scripts/update_playbook.py    --input playbook_points.json --playbook references/contract-review-playbook.md
```

> 脚本中的 `$PY` 是「Python 解释器」占位符，上面统一用 `python3` 代替。

## 用法示例

**`edits.json`**（喂给 `track_changes.py`，生成 Word 修订稿）：

```json
[
  {"action":"replace","para_idx":5,"old":"原文","new":"新文","comment":"建议改为 X。理由：Y"},
  {"action":"insert","para_idx":8,"text":"新增条款","comment":"建议补充。理由：Y"},
  {"action":"comment","para_idx":3,"comment":"此处模糊。理由：Y"}
]
```

**`review_result.json`**（喂给 `generate_report.py`，生成审核报告）：

```json
{
  "contract_name": "货物买卖合同",
  "review_date": "2026-08-09",
  "our_role": "甲方（买方）",
  "negotiation_position": "弱势方",
  "overall_risk": {"level": "中", "reason": "付款节点偏紧，质保金比例缺失"},
  "clauses": [
    {"priority":"高","name":"标的物条款","risk":"规格描述模糊","suggestion":"明确型号/数量单位","reason":"易引发交付争议"},
    {"priority":"中","name":"违约责任","risk":"仅单向约束我方","suggestion":"改为双向约定","reason":"对方违约时无救济"}
  ],
  "todos": ["向业务确认付款节点","向业务确认质保金比例"]
}
```

## 依赖

- Python 3.9+
- `python-docx`：`pip install python-docx`

## 约束

- 合同输入仅 `.docx`；修订稿/报告输出 `.docx`。
- 修订标记必须是真实 OOXML 元素，禁止 mock。
- 知识库原有内容只许追加、不许改删。
- 身份证号等敏感信息由用户本人填写，skill 不落盘。
