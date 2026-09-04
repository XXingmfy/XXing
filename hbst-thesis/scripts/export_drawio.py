#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_drawio.py — 把 .drawio 可编辑图源批量导出为 PNG（论文插图）与 .vsdx（Visio 可编辑/交审）
（hbst-thesis skill，配合 scripts/gen_diagram.py 使用）

用法:
  python export_drawio.py figs/                       # 导出目录下所有 .drawio
  python export_drawio.py figs/fig3_1.drawio          # 导出单个
  python export_drawio.py figs/ --png-dpi 200 --skip-vsdx
  python export_drawio.py --find                      # 只探测 draw.io 位置

自动探测 draw.io 顺序：
  1) 环境变量 DRAWIO_BIN
  2) PATH 中的 drawio / draw.io
  3) 常见安装路径（Windows/macOS/Linux）
  4) 桌面版 GUI 可执行文件（Windows: draw.io.exe 路径列表）

若只找到桌面版（无 CLI），脚本给出"在 draw.io 中打开 → File → Export as"的逐文件指引。
导出后建议把 .drawio/.png/.vsdx 同名三件套登记进 figure-registry.yaml（见 references/07-diagram-standards.md）。
"""
import argparse
import datetime
import os
import shutil
import subprocess
import sys

# 本机已知的 draw.io 安装位置（用户指定）+ 常规位置
KNOWN_DRAWIO = [
    r"D:\Tools_Home\drawio\draw.io\draw.io.exe",   # ← 用户本机实际安装路径
    r"C:\Program Files\draw.io\draw.io.exe",
    r"C:\Program Files (x86)\draw.io\draw.io.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\draw.io\draw.io.exe"),
    r"C:\Program Files\draw.io\drawio.exe",
    "/Applications/draw.io.app/Contents/MacOS/draw.io",
    "/Applications/Draw.io.app/Contents/MacOS/draw.io",
    "/snap/bin/drawio",
    "/usr/bin/drawio",
    "/usr/local/bin/drawio",
]
# 论文图默认输出根目录（用户指定：每次新建一个对应子文件夹）
DEFAULT_OUTPUT_ROOT = r"D:\Tools_Home\drawio\Load"


def find_drawio():
    env = os.environ.get("DRAWIO_BIN")
    if env and os.path.exists(env):
        return env
    for name in ("drawio", "draw.io"):
        p = shutil.which(name)
        if p:
            return p
    for c in KNOWN_DRAWIO:
        if os.path.exists(c):
            return c
    return None


def default_output_dir(tag=None):
    """返回 D:\\Tools_Home\\drawio\\Load\\<YYYYMMDD_HHMMSS[_tag]> 并创建。"""
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    name = f"{ts}" + (f"_{tag}" if tag else "")
    d = os.path.join(DEFAULT_OUTPUT_ROOT, name)
    os.makedirs(d, exist_ok=True)
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("target", nargs="?", help=".drawio 文件或包含 .drawio 的目录")
    ap.add_argument("--find", action="store_true", help="仅探测 draw.io 位置")
    ap.add_argument("--png-dpi", type=int, default=200)
    ap.add_argument("--skip-vsdx", action="store_true", help="不导出 .vsdx")
    ap.add_argument("--out-dir", default=None, help="导出到指定目录（并归档 .drawio 源）")
    ap.add_argument("--to-load", action="store_true", help="自动导出到 D:\\Tools_Home\\drawio\\Load\\<时间戳>\\ 新文件夹")
    ap.add_argument("--tag", default="", help="与 --to-load 搭配，为文件夹加后缀名（如论文题目标识）")
    args = ap.parse_args()

    bin_path = find_drawio()
    if args.find:
        print(bin_path or "未找到 draw.io。请安装：https://github.com/jgraph/drawio-desktop/releases")
        return 0
    if not args.target:
        ap.error("需要提供目标文件/目录（或用 --find）")
    if not bin_path:
        print("未找到 draw.io。请先安装：https://github.com/jgraph/drawio-desktop/releases")
        print("本机缺 draw.io 时，.drawio 仍是可编辑源文件，但无法自动导出 PNG/VSDX；")
        print("可在任意装有 draw.io 的机器打开导出，或在桌面版中：File → Export as → PNG / VSDX。")
        return 2

    # 默认输出：目标目录内同名导出；指定 --out-dir 时导出到该目录（并复制 .drawio 源过去）
    out_dir = None
    if args.out_dir:
        out_dir = os.path.abspath(args.out_dir)
        os.makedirs(out_dir, exist_ok=True)
    # --to-load: 自动输出到 D:\Tools_Home\drawio\Load\<时间戳[_tag]>\ 并复制源
    if args.to_load:
        out_dir = default_output_dir(args.tag)
    if os.path.isdir(args.target):
        files = sorted(f for f in os.listdir(args.target) if f.lower().endswith(".drawio"))
    else:
        files = [os.path.basename(args.target)]
        args.target = os.path.dirname(os.path.abspath(args.target)) or "."
    if not files:
        print(f"{args.target} 下没有 .drawio 文件")
        return 1

    ok, fail = 0, 0
    for f in files:
        src = os.path.join(args.target, f)
        stem = os.path.splitext(f)[0]
        # 目标目录：未指定则与源同目录
        dst_dir = out_dir or args.target
        dst_src = os.path.join(dst_dir, f)
        if out_dir:
            import shutil as _sh
            _sh.copy2(src, dst_src)  # 源文件一并归档到输出目录
            src = dst_src
        outs = [f]
        png = os.path.join(dst_dir, stem + ".png")
        cmd = [bin_path, "-x", "-f", "png", "--scale", "2" if args.png_dpi >= 200 else "1",
               "-o", png, src]
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode == 0 and os.path.exists(png):
            outs.append(os.path.basename(png)); ok += 1
        else:
            print(f"PNG 导出失败 {f}: {r.stderr[-300:] if r.stderr else ''}")
            fail += 1
            continue
        if not args.skip_vsdx:
            vsdx = os.path.join(dst_dir, stem + ".vsdx")
            cmd2 = [bin_path, "-x", "-f", "vsdx", "-o", vsdx, src]
            r2 = subprocess.run(cmd2, capture_output=True, text=True)
            if r2.returncode == 0 and os.path.exists(vsdx):
                outs.append(os.path.basename(vsdx))
            else:
                print(f"VSDX 导出失败 {f}（可后续手动导出）: {r2.stderr[-200:] if r2.stderr else ''}")
        print(f"[export_drawio] {f} -> {', '.join(outs)} (于 {dst_dir})")
    print(f"\n完成：成功 {ok} / 失败 {fail}。输出目录：{out_dir or args.target}")
    print("提醒：.drawio/.png/.vsdx 三件套登记进 figure-registry.yaml；docx 插图用 PNG，交审/留档用 .vsdx。")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
