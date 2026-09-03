#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_docx_structure.py — 阶段三：docx 结构闸门（hbst-thesis skill）

直接读 .docx 的 OOXML，检查结构与格式硬指标，输出 JSON + 退出码。
定位：把 writing-standards §9 与 docx-production §8 中"可机器判定"的部分落成确定性检查。

检查项:
  D1 分节:        节数 >= 3；正文所在节 pgNumType 为 decimal(start=1)（阿拉伯页码）
  D2 标题样式:     存在使用标题1样式的段落（章标题）；正文标题不落在纯 Normal 加粗冒充（抽样提示）
  D3 目录域:       文档含 TOC 指令（TOC \\o）；"目录"标题存在
  D4 页脚页码域:   至少一个 footer 含 PAGE 域
  D5 三线表:       表级 tblBorders 仅 top/bottom（或 top+bottom+insideH=nil），无竖线；
                   不得出现 Table Grid 网格样式残留导致 insideV/竖线可见（启发式）
  D6 续表表头:     出现"续表"文本时，附近表应含 w:tblHeader 行（启发式: 表头行 tblHeader）
  D7 声明页隔离:   文本含"原创性声明"与"版权使用授权书"且不同页（按分页符/分节粗略判断——给人工）

用法:
  python check_docx_structure.py <paper.docx> [--json out.json] [--expect-sections 6]
依赖: 标准库 zipfile/xml；不需要 python-docx。
"""
import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(tag):
    return f"{{{W}}}{tag}"


def extract_docxml(path):
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    doc = ET.fromstring(z.read("word/document.xml"))
    footers = []
    for n in sorted(names):
        if re.match(r"word/footer\d+\.xml$", n):
            footers.append(ET.fromstring(z.read(n)))
    return z, doc, footers


def para_text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def para_style(p):
    pPr = p.find(q("pPr"))
    if pPr is None:
        return None
    ps = pPr.find(q("pStyle"))
    return ps.get(q("val")) if ps is not None else None


def has_toc_field(doc):
    for it in doc.iter(q("instrText")):
        if it.text and "TOC" in it.text and "\\o" in it.text:
            return True
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--json", default=None)
    ap.add_argument("--expect-sections", type=int, default=6)
    args = ap.parse_args()

    issues = []

    def add(code, level, msg):
        issues.append({"id": code, "level": level, "message": msg})

    try:
        z, doc, footers = extract_docxml(args.input)
    except Exception as e:
        print(json.dumps({"error": f"不是可读 OOXML .docx: {e}"}, ensure_ascii=False))
        sys.exit(2)

    body = doc.find(q("body"))
    # 分节：段落级 sectPr + body 级 sectPr
    secs = list(doc.iter(q("sectPr")))
    n_sec = len(secs)
    if n_sec < 3:
        add("D1", "FAIL", f"分节数 {n_sec} < 3（学校模板模型通常 6 节）")
    else:
        add("D1", "INFO", f"分节数 {n_sec}（模板模型通常 6 节；少于预期请人工确认）")
    # 正文节（最后一个 sectPr 通常是正文~附录）页码格式
    last = secs[-1] if secs else None
    if last is not None:
        pn = last.find(q("pgNumType"))
        if pn is not None:
            fmt = pn.get(q("fmt"))
            start = pn.get(q("start"))
            add("D1", "INFO", f"末节 pgNumType fmt={fmt} start={start}（期望 decimal start=1）")
            if fmt and fmt != "decimal":
                add("D1", "WARN", f"末节页码格式 {fmt} 非 decimal，请确认正文是否阿拉伯页码")
        else:
            add("D1", "WARN", "末节无 pgNumType，页码格式未显式声明（可能沿用前节）")
    else:
        add("D1", "FAIL", "未找到任何 sectPr")

    # 标题样式抽样
    h1_count = 0
    h2_count = 0
    for p in doc.iter(q("p")):
        s = para_style(p)
        if s in ("1", "Heading1", "heading 1"):
            h1_count += 1
        elif s in ("2", "Heading2", "heading 2"):
            h2_count += 1
    if h1_count == 0:
        add("D2", "FAIL", "未找到使用标题1样式的段落（章标题必须是内置标题样式）")
    else:
        add("D2", "INFO", f"标题1样式段落 {h1_count} 个、标题2样式段落 {h2_count} 个")

    # 目录域
    if has_toc_field(doc):
        add("D3", "INFO", "检测到 TOC 域（自动目录）")
    else:
        add("D3", "FAIL", "未检测到 TOC 域（目录应为自动目录域 TOC \\o \\\"1-3\\\"，非手打）")
    has_toc_heading = any("目录" in para_text(p) or "目 录" in para_text(p) for p in doc.iter(q("p")))
    if not has_toc_heading:
        add("D3", "WARN", "未找到“目录”标题文本")

    # 页脚 PAGE 域
    footer_has_page = False
    for f in footers:
        for it in f.iter(q("instrText")):
            if it.text and "PAGE" in it.text:
                footer_has_page = True
    if footer_has_page:
        add("D4", "INFO", "页脚存在 PAGE 域（页码为域，非手打）")
    else:
        add("D4", "WARN", "未在页脚找到 PAGE 域——请确认页码是否手打（模板要求域）")

    # 三线表
    tbls = list(doc.iter(q("tbl")))
    bad_tbl = []
    for i, tb in enumerate(tbls):
        tblPr = tb.find(q("tblPr"))
        if tblPr is None:
            continue
        borders = tblPr.find(q("tblBorders"))
        if borders is None:
            bad_tbl.append((i, "无表级边框定义"))
            continue
        side_ok = True
        for side in ("left", "right", "insideV"):
            b = borders.find(q(side))
            if b is not None and b.get(q("val")) not in (None, "nil", "none"):
                side_ok = False
        # insideH 若为 single 且非表头行单独定义，判为网格残留风险
        ih = borders.find(q("insideH"))
        if ih is not None and ih.get(q("val")) in ("single", "double", "dashed"):
            side_ok = False
        if not side_ok:
            bad_tbl.append((i, "存在竖线或内部横线（非三线表）"))
    if bad_tbl:
        add("D5", "FAIL", f"{len(bad_tbl)} 张表疑似非三线表: {bad_tbl[:8]}")
    else:
        add("D5", "INFO", f"{len(tbls)} 张表表级边框均为三线表结构（top/bottom 无竖线）")

    # 续表表头
    text_all = "".join(t.text or "" for t in doc.iter(q("t")))
    has_cont = "续表" in text_all
    header_rows = 0
    for tb in tbls:
        for tr in tb.findall(q("tr"))[:1]:
            trPr = tr.find(q("trPr"))
            if trPr is not None and trPr.find(q("tblHeader")) is not None:
                header_rows += 1
    if has_cont:
        if header_rows == 0:
            add("D6", "WARN", "存在“续表”标注但未发现任何 tblHeader 重复表头行（跨页表应重复表头）")
        else:
            add("D6", "INFO", f"存在续表标注，检测到 {header_rows} 个表头重复行")
    else:
        add("D6", "INFO", "无续表标注（无跨页长表或无需续标）")

    # D7 声明/授权书同页提示
    p_orig = text_all.find("原创性声明")
    p_auth = text_all.find("版权使用授权书")
    if p_orig >= 0 and p_auth >= 0:
        add("D7", "WARN", "检测到“原创性声明”与“版权使用授权书”文本；请人工确认二者分属不同页（writing-standards §4.2）")
    else:
        add("D7", "INFO", "未同时检测到声明与授权书标题文本（可能用图片/其他表述，人工确认）")

    failed = [i for i in issues if i["level"] == "FAIL"]
    report = {
        "sections": n_sec,
        "heading1_count": h1_count,
        "heading2_count": h2_count,
        "tables": len(tbls),
        "footer_page_field": footer_has_page,
        "toc_field": has_toc_field(doc),
        "issues": issues,
        "failed_count": len(failed),
        "passed": len(failed) == 0,
        "note": "D5/D6 为启发式检查；边框继承、分页效果等最终以 PDF 版面人工核对（writing-standards §10）为准。",
    }
    out = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            f.write(out)
    print(out)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
