#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_diagram.py — 生成可编辑结构图源（draw.io 格式 .drawio）（hbst-thesis skill）

把结构化图描述生成 draw.io 可编辑源，供 draw.io（免费）打开并导出 PNG / 另存 Visio .vsdx。
满足 07-diagram-standards：结构图必须出自可编辑矢量源，禁止 matplotlib 直出 PNG 冒充。

两类用法：
1) 便捷生成器（论文写作 agent 常用）：
   python gen_diagram.py usecase   --out figs/fig3_1_usecase.drawio --actors "学生,管理员" --usecases "登录,请假申请,报修,审批"
   python gen_diagram.py flow      --out figs/fig4_1.drawio --steps "开始|提交申请|审核通过?|通过:进入下一流程;不通过:退回|结束"
   python gen_diagram.py arch      --out figs/fig4_2.drawio --layers "表现层:小程序,后台;服务层:SpringBoot;数据层:MySQL"
2) JSON 精确描述（节点/边/坐标全自定义，给复杂图）：
   python gen_diagram.py json --in graph.json --out figs/x.drawio

生成后导出（需本机装有 draw.io）：
   drawio -x -f png  -o x.png   x.drawio
   drawio -x -f vsdx -o x.vsdx  x.drawio   # 可编辑 Visio 文件

.drawio 坐标单位 px，y 向下；节点 kinds: box(矩形) / ellipse(椭圆,用例/起止) / rhom(菱形,判断) / actor(参与者)。
"""
import argparse
import html
import json
import os
import re
import sys
import xml.etree.ElementTree as ET

# draw.io (.drawio) 文件为无命名空间 XML；不要给 mxCell 等加 namespace，否则 draw.io 打不开。
TAG = "mxCell"
GEO = "mxGeometry"
POINT = "mxPoint"


def _cell(tag, **attrs):
    el = ET.Element(tag)
    for k, v in attrs.items():
        if v is not None:
            el.set(k, str(v))
    return el


def _mx_cell(node_id, parent, value, style, vertex, x, y, w, h):
    cell = _cell("mxCell", id=node_id, parent=parent, value=value or None,
                 style=style, vertex="1" if vertex else "0", edge="0" if vertex else None)
    geo = ET.SubElement(cell, "mxGeometry")
    if vertex:
        geo.set("x", str(x)); geo.set("y", str(y))
        geo.set("width", str(w)); geo.set("height", str(h))
    return cell


def _mx_edge(edge_id, parent, src, dst, label, style):
    # source/target 是 mxCell 的 XML 属性，值指向节点 cell 的 id；没有 source/target 会画成浮动线
    cell = _cell("mxCell", id=edge_id, parent=parent, source=src, target=dst,
                 value=label or None, style=style, vertex="0", edge="1")
    geo = ET.SubElement(cell, "mxGeometry")
    geo.set("relative", "1"); geo.set("as", "geometry")
    return cell


def build_drawio_xml(cells):
    """cells: list of mxCell 元素（无 namespace）。返回 .drawio 完整文本。"""
    mxfile = ET.Element("mxfile", {"host": "app.diagrams.net", "agent": "hbst-thesis", "version": "21.0.0"})
    diagram = ET.SubElement(mxfile, "diagram", {"id": "d1", "name": "Page-1"})
    model = ET.SubElement(diagram, "mxGraphModel")
    root = ET.SubElement(model, "root")
    layer0 = _cell("mxCell", id="0")
    layer1 = _cell("mxCell", id="1", parent="0")
    root.append(layer0)
    root.append(layer1)
    for c in cells:
        root.append(c)
    ET.indent(mxfile, space="  ")
    return ET.tostring(mxfile, encoding="unicode")


def txt(v):
    return html.escape(str(v), quote=False)


# ---------- 便捷生成器 ----------
def _auto_layout(n, start_x=40, start_y=40, dx=180, dy=90, cols=None):
    """一维/网格自动排布。返回 [(x,y)]。"""
    cols = cols or max(1, int((900 - start_x) / dx))
    out = []
    for i in range(n):
        r, c = divmod(i, cols)
        out.append((start_x + c * dx, start_y + r * dy))
    return out


def gen_usecase(out, actors, usecases):
    """参与者(左) + 系统边界框 + 用例椭圆(中)。actors/usecases 为列表。"""
    cells = []
    nid = 1
    # 边界框
    box_h = max(140, len(usecases) * 80 + 40)
    cells.append(_mx_cell(nid, 1, "系统", "rounded=1;whiteSpace=wrap;html=1;dashed=1;", True, 240, 40, 520, box_h))
    boundary = nid; nid += 1
    # 参与者放左侧一列
    actor_ids = []
    ay = 90
    for a in actors:
        cells.append(_mx_cell(nid, 1, txt(a), "shape=actor;whiteSpace=wrap;html=1;", True, 40, ay, 60, 90))
        actor_ids.append(nid)
        nid += 1
        ay += 120
    # 用例椭圆（boundary 内，两列排布；坐标相对 boundary）
    uc_ids = []
    rows = (len(usecases) + 1) // 2
    for i, u in enumerate(usecases):
        col = i // rows
        r = i % rows
        x = 30 + col * 230
        y = 30 + r * 90
        cells.append(_mx_cell(nid, boundary, txt(u), "ellipse;whiteSpace=wrap;html=1;", True, x, y, 200, 60))
        uc_ids.append(nid); nid += 1
    # 参与者 -> 全部用例连线（显式 source/target）
    for ai in actor_ids:
        for ui in uc_ids:
            cells.append(_mx_edge(nid, 1, ai, ui, "", "edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;endArrow=block;"))
            nid += 1
    _write(out, cells)


def gen_flow(out, steps):
    """steps: list of (text, kind) kind in box|rhom|ellipse，或 "文本|类型" 字符串列表。"""
    cells = []
    parsed = []
    for s in steps:
        if isinstance(s, str):
            if "|" in s:
                t, k = s.split("|", 1)
            else:
                t, k = s, "box"
            parsed.append((t.strip(), k.strip()))
        else:
            parsed.append(s)
    nid = 1
    x, y, w, h = 40, 40, 180, 54
    prev = None
    style_map = {"box": "rounded=1;whiteSpace=wrap;html=1;", "rhom": "rhombus;whiteSpace=wrap;html=1;",
                 "ellipse": "ellipse;whiteSpace=wrap;html=1;", "actor": "shape=actor;whiteSpace=wrap;html=1;"}
    ids = []
    for i, (t, k) in enumerate(parsed):
        st = style_map.get(k, style_map["box"])
        hh = h
        if k == "rhom":
            hh = 90
        cells.append(_mx_cell(nid, 1, txt(t), st, True, x, y + i * (hh + 40), w, hh))
        ids.append(nid)
        nid += 1
    for i in range(1, len(ids)):
        cells.append(_mx_edge(nid, 1, ids[i - 1], ids[i], "", "edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;endArrow=block;"))
        nid += 1
    _write(out, cells)


def gen_arch(out, layers):
    """layers: list of "层名: 项1,项2,..." 或 (层名, [项])。自上而下分层。"""
    cells = []
    nid = 1
    parsed = []
    for l in layers:
        if isinstance(l, str) and ":" in l:
            name, items = l.split(":", 1)
            parsed.append((name.strip(), [i.strip() for i in items.split(",") if i.strip()]))
        else:
            name, items = l
            parsed.append((name, items))
    y = 30
    width = 820
    x0 = 40
    per_layer_h = 130
    layer_first_item = []  # 每层第一个模块的 cell id（用于层间连线）
    for li, (name, items) in enumerate(parsed):
        ly = y + li * per_layer_h
        # 层容器
        cells.append(_mx_cell(nid, 1, txt(name), "rounded=0;whiteSpace=wrap;html=1;verticalAlign=top;fontStyle=1;", True, x0, ly, width, 40))
        layer_cell = nid; nid += 1
        # 层内模块横排
        first_item = None
        if items:
            iw = min(160, (width - 20 - 10 * (len(items) - 1)) // max(1, len(items)))
            for ii, it in enumerate(items):
                ix = x0 + 10 + ii * (iw + 10)
                cells.append(_mx_cell(nid, layer_cell, txt(it), "rounded=1;whiteSpace=wrap;html=1;", True, ix, ly + 45, iw, 50))
                if first_item is None:
                    first_item = nid
                nid += 1
        layer_first_item.append(first_item)
    # 层间箭头：上层首模块 -> 下层首模块（表示调用/数据流），上下相邻层之间
    for i in range(1, len(layer_first_item)):
        if layer_first_item[i - 1] and layer_first_item[i]:
            cells.append(_mx_edge(nid, 1, layer_first_item[i - 1], layer_first_item[i], "",
                                  "edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;endArrow=block;exitX=0.5;exitY=1;entryX=0.5;entryY=0;"))
            nid += 1
    _write(out, cells)


def gen_er(out, entities, relations):
    """entities: [(名称, 颜色/标注)] ; relations: [(from, to, label)]，E-R 实体框。"""
    cells = []
    nid = 1
    pos = _auto_layout(len(entities), start_x=60, start_y=60, dx=300, dy=220, cols=2)
    ent_ids = {}
    for (name, _), (x, y) in zip(entities, pos):
        cells.append(_mx_cell(nid, 1, txt(name), "rounded=1;whiteSpace=wrap;html=1;fontStyle=1;", True, x, y, 220, 90))
        ent_ids[name] = nid; nid += 1
    for frm, to, label in relations:
        if frm in ent_ids and to in ent_ids:
            cells.append(_mx_edge(nid, 1, ent_ids[frm], ent_ids[to], txt(label), "endArrow=block;startArrow=none;html=1;"))
            nid += 1
    _write(out, cells)


def gen_json(out, data):
    """JSON: {nodes:[{id,label,kind,x,y,w,h,parent}], edges:[{from,to,label}]}"""
    cells = []
    nid = 1000
    ids = {}
    style_map = {"box": "rounded=1;whiteSpace=wrap;html=1;", "ellipse": "ellipse;whiteSpace=wrap;html=1;",
                 "rhom": "rhombus;whiteSpace=wrap;html=1;", "actor": "shape=actor;whiteSpace=wrap;html=1;",
                 "plain": "rounded=0;whiteSpace=wrap;html=1;"}
    for n in data.get("nodes", []):
        nd = nid; nid += 1
        ids[n.get("id", nd)] = nd
        st = style_map.get(n.get("kind", "box"), style_map["box"])
        cells.append(_mx_cell(nd, n.get("parent", 1), txt(n.get("label", "")), st, True,
                              n.get("x", 0), n.get("y", 0), n.get("w", 120), n.get("h", 50)))
    for e in data.get("edges", []):
        src = ids.get(e.get("from")); dst = ids.get(e.get("to"))
        cells.append(_mx_edge(nid, 1, src, dst, txt(e.get("label", "")),
                              e.get("style", "edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;endArrow=block;")))
        nid += 1
    _write(out, cells)


def _write(out, cells):
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    xml_text = build_drawio_xml(cells)
    with open(out, "w", encoding="utf-8") as f:
        f.write(xml_text)
    print(f"[gen_diagram] wrote {out}")
    print("[gen_diagram] 下一步: 用 draw.io 打开并导出 —— "
          f"drawio -x -f png -o {os.path.splitext(out)[0]}.png {out} ; "
          f"drawio -x -f vsdx -o {os.path.splitext(out)[0]}.vsdx {out}")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p_uc = sub.add_parser("usecase"); p_uc.add_argument("--out", required=True)
    p_uc.add_argument("--actors", required=True, help="逗号分隔参与者")
    p_uc.add_argument("--usecases", required=True, help="逗号分隔用例")
    p_fl = sub.add_parser("flow"); p_fl.add_argument("--out", required=True)
    p_fl.add_argument("--steps", required=True, help="竖排步骤，用|分隔，如 '开始|提交?|结束'；'?'结尾判为判断菱形")
    p_ar = sub.add_parser("arch"); p_ar.add_argument("--out", required=True)
    p_ar.add_argument("--layers", required=True, help="层描述，分号分隔，如 '表现层:小程序,后台;服务层:SpringBoot;数据层:MySQL'")
    p_er = sub.add_parser("er"); p_er.add_argument("--out", required=True)
    p_er.add_argument("--entities", required=True, help="逗号分隔实体名")
    p_er.add_argument("--relations", default="", help="关系描述 'A-B:标注;B-C:标注'")
    p_js = sub.add_parser("json"); p_js.add_argument("--out", required=True); p_js.add_argument("--in", dest="infile", required=True)
    args = ap.parse_args()

    if args.cmd == "usecase":
        gen_usecase(args.out, [a.strip() for a in args.actors.split(",") if a.strip()],
                    [u.strip() for u in args.usecases.split(",") if u.strip()])
    elif args.cmd == "flow":
        steps = []
        for s in args.steps.split("|"):
            s = s.strip()
            if not s:
                continue
            kind = "rhom" if s.endswith("?") else "box"
            if kind == "rhom":
                s = s[:-1]
            steps.append((s, kind))
        gen_flow(args.out, steps)
    elif args.cmd == "arch":
        gen_arch(args.out, [l.strip() for l in args.layers.split(";") if l.strip()])
    elif args.cmd == "er":
        ents = [(e.strip(), "") for e in args.entities.split(",") if e.strip()]
        rels = []
        for r in args.relations.split(";"):
            r = r.strip()
            if not r:
                continue
            if ":" in r:
                pair, lab = r.split(":", 1)
            else:
                pair, lab = r, ""
            if "-" in pair:
                a, b = pair.split("-", 1)
                rels.append((a.strip(), b.strip(), lab.strip()))
        gen_er(args.out, ents, rels)
    elif args.cmd == "json":
        with open(args.infile, "r", encoding="utf-8") as f:
            gen_json(args.out, json.load(f))


if __name__ == "__main__":
    main()
