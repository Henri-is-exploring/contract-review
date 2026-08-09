#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""track_changes.py — 对 .docx 应用 Word 修订标记(w:ins/w:del)和批注(comments.xml)。

用法:
  $PY track_changes.py --input contract.docx --edits edits.json --output revised.docx

指令 JSON 格式:
  [
    {"action":"replace","para_idx":5,"old":"原文","new":"新文","comment":"建议改为X。理由：Y"},
    {"action":"insert","para_idx":8,"text":"新增条款","comment":"建议补充。理由：Y"},
    {"action":"comment","para_idx":3,"comment":"此处模糊。理由：Y"}
  ]

死规矩:
  - w:ins 包裹插入文本,设 w:author/w:date; w:del+w:delText 标记删除
  - 批注用 comments.xml part,格式统一「建议+理由」
  - 必须真实 OOXML 元素,不模拟样式
"""
import argparse
import json
import sys
from datetime import datetime, timezone

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RT_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
AUTHOR = '法务合同审核'
INITIALS = '法'

# 修订版(revised.docx)直接发业务,不得出现内部知识库来源引用。
# 对 new/text(入合同条款) 与 comment(批注) 均做禁用词校验,命中即拒绝生成。
FORBIDDEN = [
    '审核手册', '审核要点手册', '合同审核要点手册',
    '知识库', 'playbook', '审查要点', '审核经验积累',
    '审核要点', '依据手册', '参照手册', '见手册',
]


def _check_leak(edits):
    """检查 edits 是否含内部知识库引用。命中则打印明细并 exit(1),不生成 docx。"""
    bad = []
    for i, e in enumerate(edits):
        for field in ('new', 'text', 'comment'):
            val = e.get(field)
            if not val:
                continue
            for kw in FORBIDDEN:
                if kw in val:
                    bad.append((i, e.get('action'), field, kw, val))
    if bad:
        sys.stderr.write('[拒绝生成] 修订版直接发业务,不得出现内部知识库引用。命中禁用词:\n')
        for i, action, field, kw, val in bad:
            sys.stderr.write(
                f'  edit#{i} ({action}) 字段 {field} 含「{kw}」: {val[:60]!r}\n')
        sys.stderr.write(
            '请改写 edits.json:new/text 须为纯条款文字;'
            'comment 理由只用业务与法律风险语言,不得引用审核手册/知识库来源。\n')
        sys.exit(1)


def _now():
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


def _run(text, deltext=False):
    """构造 <w:r><w:t|w:delText xml:space=preserve>text</...></w:r>"""
    r = OxmlElement('w:r')
    t = OxmlElement('w:delText' if deltext else 'w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    return r


def _ins(text, wid):
    e = OxmlElement('w:ins')
    e.set(qn('w:id'), str(wid))
    e.set(qn('w:author'), AUTHOR)
    e.set(qn('w:date'), _now())
    e.append(_run(text))
    return e


def _del(text, wid):
    e = OxmlElement('w:del')
    e.set(qn('w:id'), str(wid))
    e.set(qn('w:author'), AUTHOR)
    e.set(qn('w:date'), _now())
    e.append(_run(text, deltext=True))
    return e


def _crs(cid):
    e = OxmlElement('w:commentRangeStart'); e.set(qn('w:id'), str(cid)); return e


def _cre(cid):
    e = OxmlElement('w:commentRangeEnd'); e.set(qn('w:id'), str(cid)); return e


def _cref(cid):
    r = OxmlElement('w:r')
    ref = OxmlElement('w:commentReference')
    ref.set(qn('w:id'), str(cid))
    r.append(ref)
    return r


def _para_text(p):
    return ''.join((t.text or '') for t in p.findall('.//' + qn('w:t')))


def _strip_runs(p):
    """移除段落下除 w:pPr 外的所有子元素"""
    for c in list(p):
        if c.tag != qn('w:pPr'):
            p.remove(c)


class Editor:
    def __init__(self, doc):
        self.doc = doc
        # 存底层 lxml element (CT_P),可直接调用 findall/append/addnext/find/insert
        self.paras = [p._p for p in doc.paragraphs]
        self.wid = 1000
        self.cid = 0
        self.comments = []  # [(cid, text)]

    def _wid(self):
        self.wid += 1
        return self.wid

    def _ncid(self):
        c = self.cid
        self.cid += 1
        return c

    def replace(self, edit):
        idx = edit['para_idx']; old = edit['old']; new = edit['new']
        cmt = edit.get('comment')
        p = self.paras[idx]
        full = _para_text(p)
        if old not in full:
            raise ValueError(f"[replace] para_idx={idx} 未找到 old 文本: {old[:30]!r}")
        i = full.index(old)
        before, after = full[:i], full[i + len(old):]
        _strip_runs(p)
        if before:
            p.append(_run(before))
        cid = self._ncid() if cmt else None
        if cid is not None:
            p.append(_crs(cid))
        p.append(_del(old, self._wid()))
        p.append(_ins(new, self._wid()))
        if cid is not None:
            p.append(_cre(cid))
            p.append(_cref(cid))
            self.comments.append((cid, cmt))
        if after:
            p.append(_run(after))

    def insert(self, edit):
        idx = edit['para_idx']; text = edit['text']
        cmt = edit.get('comment')
        p = self.paras[idx]
        new_p = OxmlElement('w:p')
        cid = self._ncid() if cmt else None
        if cid is not None:
            new_p.append(_crs(cid))
        new_p.append(_ins(text, self._wid()))
        if cid is not None:
            new_p.append(_cre(cid))
            new_p.append(_cref(cid))
            self.comments.append((cid, cmt))
        p.addnext(new_p)

    def comment(self, edit):
        idx = edit['para_idx']; cmt = edit['comment']
        p = self.paras[idx]
        cid = self._ncid()
        pPr = p.find(qn('w:pPr'))
        if pPr is not None:
            pPr.addnext(_crs(cid))
        else:
            p.insert(0, _crs(cid))
        p.append(_cre(cid))
        p.append(_cref(cid))
        self.comments.append((cid, cmt))

    def _comments_xml(self):
        from lxml import etree
        root = etree.Element(qn('w:comments'), nsmap={'w': W_NS})
        for cid, text in self.comments:
            c = etree.SubElement(root, qn('w:comment'))
            c.set(qn('w:id'), str(cid))
            c.set(qn('w:author'), AUTHOR)
            c.set(qn('w:date'), _now())
            c.set(qn('w:initials'), INITIALS)
            cp = etree.SubElement(c, qn('w:p'))
            cr = etree.SubElement(cp, qn('w:r'))
            ct = etree.SubElement(cr, qn('w:t'))
            ct.set(qn('xml:space'), 'preserve')
            ct.text = text
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                              standalone=True)

    def _attach(self, xml_bytes):
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        partname = PackURI('/word/comments.xml')
        part = Part(partname, CT_COMMENTS, xml_bytes, self.doc.part.package)
        self.doc.part.relate_to(part, RT_COMMENTS)

    def save(self, out):
        if self.comments:
            self._attach(self._comments_xml())
        self.doc.save(out)


def main():
    ap = argparse.ArgumentParser(description='对 docx 应用 Word 修订+批注(真实 OOXML)')
    ap.add_argument('--input', required=True, help='原始 .docx 路径')
    ap.add_argument('--edits', required=True, help='修订指令 JSON 路径')
    ap.add_argument('--output', required=True, help='输出 .docx 路径')
    args = ap.parse_args()

    with open(args.edits, encoding='utf-8') as f:
        edits = json.load(f)

    _check_leak(edits)  # 修订版发业务前,先拦内部知识库引用

    doc = Document(args.input)
    ed = Editor(doc)
    # 先处理非 insert(不增段),再反向处理 insert(保持 edits 原序)
    for e in edits:
        if e['action'] == 'replace':
            ed.replace(e)
        elif e['action'] == 'comment':
            ed.comment(e)
        elif e['action'] != 'insert':
            raise ValueError(f"未知 action: {e['action']}")
    for e in reversed([e for e in edits if e['action'] == 'insert']):
        ed.insert(e)

    ed.save(args.output)
    n_rev = sum(1 for e in edits if e['action'] in ('replace', 'insert'))
    print(f'OK -> {args.output} | 修订动作:{n_rev} 批注:{len(ed.comments)}',
          file=sys.stderr)


if __name__ == '__main__':
    main()
