#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 结果汇总.md 转为带样式的单文件 HTML（手机/电脑浏览器直接看）。"""
import html
import re
import sys

SRC = sys.argv[1] if len(sys.argv) > 1 else "考场安排输出/结果汇总.md"
OUT = sys.argv[2] if len(sys.argv) > 2 else "考场安排输出/结果汇总.html"

lines = open(SRC, encoding="utf-8").read().split("\n")

out = []
buf_table = []


def flush_table():
    global buf_table
    if not buf_table:
        return
    rows = [r for r in buf_table if not re.fullmatch(r"\|[\s\-:|]+\|", r)]
    out.append('<div class="tw"><table>')
    for i, r in enumerate(rows):
        cells = [c.strip() for c in r.strip().strip("|").split("|")]
        tag = "th" if i == 0 else "td"
        out.append("<tr>" + "".join(
            "<%s>%s</%s>" % (tag, html.escape(c), tag) for c in cells) + "</tr>")
    out.append("</table></div>")
    buf_table = []


def inline(t):
    t = html.escape(t)
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"`(.+?)`", r"<code>\1</code>", t)
    return t


for ln in lines:
    if ln.startswith("|"):
        buf_table.append(ln)
        continue
    flush_table()
    if ln.startswith("### "):
        out.append("<h3>%s</h3>" % inline(ln[4:]))
    elif ln.startswith("## "):
        out.append("<h2>%s</h2>" % inline(ln[3:]))
    elif ln.startswith("# "):
        out.append("<h1>%s</h1>" % inline(ln[2:]))
    elif ln.startswith("> "):
        out.append("<blockquote>%s</blockquote>" % inline(ln[2:]))
    elif ln.startswith("- "):
        out.append("<li>%s</li>" % inline(ln[2:]))
    elif ln.strip() == "":
        out.append("")
    else:
        out.append("<p>%s</p>" % inline(ln))
flush_table()

doc = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>2026.9 高三考场安排结果（24级）</title>
<style>
:root{--line:#e3e6ea;--ink:#1f2328;--sub:#6b7280;--brand:#2f5d8a;--hl:#fff8e1}
*{box-sizing:border-box}
body{margin:0;padding:20px 16px 60px;font:15px/1.7 -apple-system,"PingFang SC","Microsoft YaHei",sans-serif;color:var(--ink);background:#f7f8fa}
h1{font-size:22px;margin:8px 0 4px;color:var(--brand)}
h2{font-size:18px;margin:28px 0 8px;padding-left:10px;border-left:4px solid var(--brand)}
h3{font-size:16px;margin:20px 0 6px;color:var(--brand)}
blockquote{margin:8px 0;padding:8px 12px;background:var(--hl);border-left:3px solid #f0c36d;color:#6b5a2e;font-size:13.5px;border-radius:0 6px 6px 0}
code{background:#eef1f4;padding:1px 5px;border-radius:4px;font-size:13px}
li{margin:2px 0 2px 18px}
.tw{overflow-x:auto;margin:10px 0 22px;border:1px solid var(--line);border-radius:8px;background:#fff}
table{border-collapse:collapse;width:100%;font-size:12.5px;white-space:nowrap}
th,td{border:1px solid var(--line);padding:4px 8px;text-align:center}
th{background:#eef3f8;color:var(--brand);position:sticky;top:0}
th,td{white-space:nowrap}
tr:nth-child(even){background:#fafbfc}
p{color:var(--sub);font-size:13.5px}
@page{size:A4 landscape;margin:10mm}
@media print{
 body{background:#fff;padding:0;font-size:11px}
 h1{font-size:18px}h2{font-size:14px;margin:14px 0 6px}h3{font-size:12px;margin:10px 0 4px}
 .tw{border:none;margin:4px 0 10px;overflow:visible}
 table{font-size:8.5px}
 th,td{padding:1.5px 3px}
 tr{page-break-inside:avoid}
 h2{break-after:avoid}
 blockquote{font-size:10px;padding:4px 8px}
}
</style></head><body>
__CONTENT__
</body></html>"""

open(OUT, "w", encoding="utf-8").write(doc.replace("__CONTENT__", "\n".join(out)))
print("已生成 %s" % OUT)
