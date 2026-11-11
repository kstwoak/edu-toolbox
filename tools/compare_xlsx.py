import sys, openpyxl
from openpyxl.utils import get_column_letter

A = sys.argv[1]
B = sys.argv[2]
wa = openpyxl.load_workbook(A)
wb = openpyxl.load_workbook(B)

print("sheet 列表一致:", wa.sheetnames == wb.sheetnames)
if wa.sheetnames != wb.sheetnames:
    print("  A:", wa.sheetnames)
    print("  B:", wb.sheetnames)

def style_key(c):
    f = c.fill
    fg = ""
    if f is not None and f.fill_type:
        col = f.fgColor
        fg = "%s|%s|%s|%s" % (f.fill_type, col.type, col.rgb if col.type in ("rgb", "indexed") else col.theme, col.tint)
    fo = c.font
    al = c.alignment
    bd = c.border
    bkey = "".join((s.style or "") for s in (bd.left, bd.right, bd.top, bd.bottom))
    return "%s;%s,%s,%s,%s;%s,%s,%s;%s;%s" % (
        fg, fo.bold, fo.size, fo.name, (fo.color.rgb if fo.color is not None and fo.color.type == "rgb" else ""),
        al.horizontal, al.vertical, al.wrap_text, bkey, c.number_format)

total = diffs = 0
detail = []
for sn in wa.sheetnames:
    if sn not in wb.sheetnames:
        continue
    sa, sb = wa[sn], wb[sn]
    # 合并单元格
    ma = set(str(r) for r in sa.merged_cells.ranges)
    mb = set(str(r) for r in sb.merged_cells.ranges)
    if ma != mb:
        detail.append("[%s] 合并单元格不一致: 仅A=%s 仅B=%s" % (sn, sorted(ma - mb)[:5], sorted(mb - ma)[:5]))
    # 列宽
    for k in set(list(sa.column_dimensions.keys()) + list(sb.column_dimensions.keys())):
        va = sa.column_dimensions[k].width if k in sa.column_dimensions else None
        vb = sb.column_dimensions[k].width if k in sb.column_dimensions else None
        if (va or 0) != (vb or 0):
            detail.append("[%s] 列宽 %s: %s vs %s" % (sn, k, va, vb))
    # 行高
    for k in set(list(sa.row_dimensions.keys()) + list(sb.row_dimensions.keys())):
        va = sa.row_dimensions[k].height if k in sa.row_dimensions else None
        vb = sb.row_dimensions[k].height if k in sb.row_dimensions else None
        if (va or 0) != (vb or 0):
            detail.append("[%s] 行高 %s: %s vs %s" % (sn, k, va, vb))
    # 打印设置
    pa, pb = sa.page_setup, sb.page_setup
    if (pa.orientation, pa.paperSize, pa.fitToWidth, pa.fitToHeight) != (pb.orientation, pb.paperSize, pb.fitToWidth, pb.fitToHeight):
        detail.append("[%s] 打印设置: %s vs %s" % (sn, (pa.orientation, pa.paperSize, pa.fitToWidth, pa.fitToHeight),
                                                 (pb.orientation, pb.paperSize, pb.fitToWidth, pb.fitToHeight)))
    if sa.freeze_panes != sb.freeze_panes:
        detail.append("[%s] 冻结窗格: %s vs %s" % (sn, sa.freeze_panes, sb.freeze_panes))
    # 单元格
    ra, rb = sa.max_row, sb.max_row
    ca, cb = sa.max_column, sb.max_column
    if (ra, ca) != (rb, cb):
        detail.append("[%s] 尺寸: %dx%d vs %dx%d" % (sn, ra, ca, rb, cb))
    for r in range(1, max(ra, rb) + 1):
        for c in range(1, max(ca, cb) + 1):
            xa, xb = sa.cell(r, c), sb.cell(r, c)
            total += 1
            if xa.value != xb.value:
                diffs += 1
                if len(detail) < 40:
                    detail.append("[%s] %s%s 值: %r vs %r" % (sn, get_column_letter(c), r, xa.value, xb.value))
                continue
            ka, kb = style_key(xa), style_key(xb)
            if ka != kb:
                diffs += 1
                if len(detail) < 40:
                    detail.append("[%s] %s%s 样式: %s vs %s" % (sn, get_column_letter(c), r, ka, kb))

print("比对单元格: %d，差异: %d" % (total, diffs))
print("结构/样式问题 %d 条" % len(detail))
for d in detail[:40]:
    print("  ", d)
