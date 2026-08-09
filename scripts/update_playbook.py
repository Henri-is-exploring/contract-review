#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""update_playbook.py — 将审核要点回补到 playbook.md 对应章节。

用法:
  $PY update_playbook.py --input points.json --playbook references/contract-review-playbook.md

死规矩:
  - 只追加不修改/删除已有内容
  - 追加到对应章节的「#### 审核经验积累」子节(无则创建)
  - 每条格式: - [ ] 检查项（来源：YYYY-MM-DD 合同审核）
  - 去重: 归一化后与同章节已有经验项比对，重复则跳过并报告（不写入）
  - 回补后向 stdout 输出已追加条目清单,供用户确认

输入 JSON 格式:
  {
    "review_date": "2026-08-09",
    "points": [
      {"section": "（四）违约责任", "check": "违约责任须双向约定，不能只约束一方"},
      {"section": "（六）争议解决", "check": "管辖须约定具体法院/仲裁机构，禁用「有管辖权的法院」"}
    ]
  }
"""
import argparse
import json
import re
import sys
from datetime import datetime

EXP_HEADER = '#### 审核经验积累'


def _norm(text):
    """归一化检查项文本用于去重：去掉来源后缀、首尾空白与末尾标点。"""
    t = text
    # 去掉「（来源：... 合同审核）」/ "(来源:... 合同审核)" 后缀
    t = re.sub(r'（来源：[^）]*合同审核）\s*$', '', t)
    t = re.sub(r'\(来源：[^)]*合同审核\)\s*$', '', t)
    t = t.strip()
    # 去掉行首的列表标记
    t = re.sub(r'^-\s*\[\s*\]\s*', '', t)
    # 去掉末尾常见标点
    t = t.rstrip('。.；;，,')
    return t.strip()


def _existing_exp_norms(lines, start, end):
    """收集 [start,end) 内已有经验项('- [ ] ...')的归一化文本集合。"""
    norms = set()
    for i in range(start + 1, end):
        ln = lines[i].strip()
        if ln.startswith('- [ ]'):
            norms.add(_norm(ln[len('- [ ]'):].strip()))
    return norms


def _find_section_bounds(lines, section_keyword):
    """定位 ### 小节起止行号(半开区间 [start, end))。
    start: '### ...section_keyword...' 所在行
    end:   下一个 '### ' 或 '## '(同级及以上) 所在行,或文件尾
    返回 (start, end) 或 None。
    """
    start = None
    for i, ln in enumerate(lines):
        if ln.startswith('### ') and section_keyword in ln:
            start = i
            break
    if start is None:
        return None
    end = len(lines)
    for j in range(start + 1, len(lines)):
        ln = lines[j]
        if ln.startswith('### '):
            end = j
            break
        # ## 级边界: 但「#### 审核经验积累」属本小节内容,不算边界
        if ln.startswith('## ') and ln.strip() != EXP_HEADER:
            end = j
            break
    return (start, end)


def _find_exp_block(lines, start, end):
    """在小节 [start,end) 内查找已有的「#### 审核经验积累」标题行号。
    返回该标题行号(追加点为其后第一个非经验内容行),无则返回 None。
    """
    for i in range(start + 1, end):
        if lines[i].strip() == EXP_HEADER:
            return i
    return None


def _insert_point(lines, idx, text):
    """在 idx 处插入一行 text(lines 在此处会被插入)"""
    lines.insert(idx, text)


def update(playbook_path, points, review_date):
    with open(playbook_path, 'r', encoding='utf-8') as f:
        content = f.read()
    # 保留原始结尾换行状态
    had_trailing_newline = content.endswith('\n')
    lines = content.split('\n')

    appended = []   # (section, line_text)
    skipped = []    # (section, line_text)  —— 去重跳过的
    # 按 section 聚合,避免多次定位
    from collections import defaultdict
    by_section = defaultdict(list)
    for p in points:
        by_section[p['section']].append(p['check'])

    # 从后往前插入,避免行号错乱(不同 section 互不影响,但同 section 多条需顺序)
    section_items = list(by_section.items())
    for section, checks in reversed(section_items):
        bounds = _find_section_bounds(lines, section)
        if bounds is None:
            # 找不到对应章节:跳过并记录,不阻断流程
            for chk in checks:
                appended.append((section, f'[未找到章节,未写入] {chk}'))
            continue
        start, end = bounds
        existing = _existing_exp_norms(lines, start, end)
        exp_idx = _find_exp_block(lines, start, end)

        # 本轮去重：与章节内已有项 + 本轮已加入项比对，重复则跳过
        to_add = []
        for chk in checks:
            n = _norm(chk)
            if n in existing or any(_norm(c) == n for c in to_add):
                skipped.append((section, f'- [ ] {chk}（来源：{review_date} 合同审核）'))
                continue
            to_add.append(chk)
            existing.add(n)

        if not to_add:
            continue

        if exp_idx is None:
            # 创建经验块:在小节末尾(end 之前)插入标题 + 检查项
            insert_at = end
            # 跳过末尾空行
            while insert_at > start + 1 and lines[insert_at - 1].strip() == '':
                insert_at -= 1
            block = ['', EXP_HEADER]
            for chk in to_add:
                block.append(f'- [ ] {chk}（来源：{review_date} 合同审核）')
            for offset, txt in enumerate(reversed(block)):
                lines.insert(insert_at, txt)
            for chk in to_add:
                appended.append((section, f'- [ ] {chk}（来源：{review_date} 合同审核）'))
        else:
            # 已有经验块:追加到该块的末尾(下一个非 '- [ ]' 行之前)
            insert_at = end
            for k in range(exp_idx + 1, end):
                if not lines[k].startswith('- [ ]'):
                    insert_at = k
                    break
            # 倒序插入保持顺序
            for chk in reversed(to_add):
                lines.insert(insert_at, f'- [ ] {chk}（来源：{review_date} 合同审核）')
            for chk in to_add:
                appended.append((section, f'- [ ] {chk}（来源：{review_date} 合同审核）'))

    new_content = '\n'.join(lines)
    if had_trailing_newline and not new_content.endswith('\n'):
        new_content += '\n'
    with open(playbook_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    # 输出已追加条目清单供用户确认(到 stdout)
    print('===== 已回补审核要点（待用户确认）=====')
    for section, txt in appended:
        print(f'[{section}] {txt}')
    written = [a for a in appended if not a[1].startswith("[未找到")]
    print(f'共 {len(written)} 条已写入；去重跳过 {len(skipped)} 条。')
    for section, txt in skipped:
        print(f'[去重跳过][{section}] {txt}')
    return appended


def main():
    ap = argparse.ArgumentParser(description='将审核要点回补到 playbook.md')
    ap.add_argument('--input', required=True, help='审核要点 JSON 路径')
    ap.add_argument('--playbook', required=True, help='playbook.md 路径')
    args = ap.parse_args()
    with open(args.input, encoding='utf-8') as f:
        data = json.load(f)
    review_date = data.get('review_date', datetime.now().strftime('%Y-%m-%d'))
    points = data.get('points', [])
    update(args.playbook, points, review_date)


if __name__ == '__main__':
    main()
