from pathlib import Path
from math import sin
from random import Random

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "泛FOF型多资产理财产品_渠道路演报告_V1.1_跨平台字体版.docx"
ASSET_DIR = ROOT / ".roadshow_assets"
SOURCE_MEDIA = ROOT / ".source_doc_media"
ASSET_DIR.mkdir(exist_ok=True)

BLUE = "163A5F"
MID_BLUE = "2E74B5"
LIGHT_BLUE = "EAF2F8"
PALE_BLUE = "F4F8FB"
TEAL = "2A7F8E"
GOLD = "B8872D"
PALE_GOLD = "FFF7E6"
INK = "1D2733"
GRAY = "5F6B76"
LIGHT_GRAY = "F2F4F7"
BORDER = "CBD5DF"
WHITE = "FFFFFF"
RED = "9B1C1C"

FONT_CN = "Microsoft YaHei"
FONT_LATIN = "Arial"


def rgb(hex_color):
    return RGBColor.from_string(hex_color)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for tag, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color=BORDER, size=6):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:color"), color)


def set_table_geometry(table, widths_dxa, indent=120):
    total = sum(widths_dxa)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_pr = table._tbl.tblPr

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")

    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), str(indent))
    tbl_ind.set(qn("w:type"), "dxa")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)

    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        tr_pr.append(cant_split)
        for idx, cell in enumerate(row.cells):
            width = widths_dxa[min(idx, len(widths_dxa) - 1)]
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            cell.width = Inches(width / 1440)
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_run_font(run, size=11, bold=None, color=INK, italic=None, name=FONT_LATIN):
    run.font.name = name
    run._element.get_or_add_rPr()
    run._element.rPr.rFonts.set(qn("w:ascii"), name)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), name)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)
    run.font.size = Pt(size)
    run.font.color.rgb = rgb(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def add_field(paragraph, instr):
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = instr
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_char1, instr_text, fld_char2])


def set_repeat_keep(paragraph, keep_next=False, keep_lines=True):
    paragraph.paragraph_format.keep_with_next = keep_next
    paragraph.paragraph_format.keep_together = keep_lines


def add_para(doc, text="", size=11, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             before=0, after=8, line=1.333, italic=False, style=None):
    p = doc.add_paragraph(style=style)
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold, color=color, italic=italic)
    return p


def add_bullets(doc, items, level=0):
    paragraphs = []
    for item in items:
        p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
        p.paragraph_format.left_indent = Inches(0.375 + 0.25 * level)
        p.paragraph_format.first_line_indent = Inches(-0.194)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.208
        r = p.add_run(item)
        set_run_font(r, size=10.6, color=INK)
        paragraphs.append(p)
    return paragraphs


def add_numbered(doc, items):
    for idx, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.375)
        p.paragraph_format.first_line_indent = Inches(-0.194)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.208
        r = p.add_run(f"{idx}.  {item}")
        set_run_font(r, size=10.6, color=INK)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    return p


def add_table(doc, headers, rows, widths, header_fill=BLUE, font_size=9.2,
              first_col_bold=False, alignments=None):
    table = doc.add_table(rows=1, cols=len(headers))
    set_table_geometry(table, widths)
    set_table_borders(table)
    hdr = table.rows[0]
    repeat_table_header(hdr)
    for j, value in enumerate(headers):
        cell = hdr.cells[j]
        set_cell_shading(cell, header_fill)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(str(value))
        set_run_font(r, size=9.2, bold=True, color=WHITE)
    for i, row in enumerate(rows):
        cells = table.add_row().cells
        for j, value in enumerate(row):
            if i % 2 == 1:
                set_cell_shading(cells[j], PALE_BLUE)
            p = cells[j].paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.12
            if alignments:
                p.alignment = alignments[j]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(str(value))
            set_run_font(r, size=font_size, bold=(first_col_bold and j == 0), color=INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_source(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(f"资料来源：{text}")
    set_run_font(r, size=8, color=GRAY, italic=True)
    return p


def add_callout(doc, title, body, fill=LIGHT_BLUE, accent=TEAL):
    table = doc.add_table(rows=1, cols=1)
    set_table_geometry(table, [9360])
    set_table_borders(table, color=accent, size=10)
    cell = table.cell(0, 0)
    set_cell_shading(cell, fill)
    set_cell_margins(cell, top=140, bottom=140, start=180, end=180)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title)
    set_run_font(r, size=11, bold=True, color=accent)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.2
    r2 = p2.add_run(body)
    set_run_font(r2, size=10.2, color=INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_placeholder(doc, code, title, purpose, method, output):
    table = doc.add_table(rows=1, cols=1)
    set_table_geometry(table, [9360])
    set_table_borders(table, color=GOLD, size=10)
    cell = table.cell(0, 0)
    set_cell_shading(cell, PALE_GOLD)
    set_cell_margins(cell, top=140, bottom=140, start=180, end=180)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(f"【待测算图表 {code}】{title}")
    set_run_font(r, size=11, bold=True, color=GOLD)
    for label, value in (("用途", purpose), ("测算口径", method), ("建议输出", output)):
        pp = cell.add_paragraph()
        pp.paragraph_format.space_after = Pt(2)
        pp.paragraph_format.line_spacing = 1.12
        rr = pp.add_run(f"{label}：")
        set_run_font(rr, size=9.5, bold=True, color=BLUE)
        rr2 = pp.add_run(value)
        set_run_font(rr2, size=9.5, color=INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_asset_slot(doc, code, title, placement, standard, output):
    table = doc.add_table(rows=1, cols=1)
    set_table_geometry(table, [9360])
    set_table_borders(table, color=TEAL, size=10)
    cell = table.cell(0, 0)
    set_cell_shading(cell, PALE_BLUE)
    set_cell_margins(cell, top=140, bottom=140, start=180, end=180)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(f"【图表素材位 {code}】{title}")
    set_run_font(r, size=11, bold=True, color=TEAL)
    for label, value in (("建议位置", placement), ("统一口径", standard), ("建议输出", output)):
        pp = cell.add_paragraph()
        pp.paragraph_format.space_after = Pt(2)
        pp.paragraph_format.line_spacing = 1.12
        rr = pp.add_run(f"{label}：")
        set_run_font(rr, size=9.5, bold=True, color=BLUE)
        rr2 = pp.add_run(value)
        set_run_font(rr2, size=9.5, color=INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_figure(doc, path, caption, source=None, width=6.35):
    if not Path(path).exists():
        return False
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(4)
    p.add_run().add_picture(str(path), width=Inches(width))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(2)
    r = cap.add_run(caption)
    set_run_font(r, size=8.8, color=GRAY)
    if source:
        add_source(doc, source)
    return True


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = paragraph.add_run("第 ")
    set_run_font(r, size=8.5, color=GRAY)
    add_field(paragraph, "PAGE")
    r2 = paragraph.add_run(" 页")
    set_run_font(r2, size=8.5, color=GRAY)


def configure_styles(doc):
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)
    sec.header_distance = Inches(0.492)
    sec.footer_distance = Inches(0.492)
    sec.different_first_page_header_footer = True

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT_LATIN
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT_LATIN)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT_LATIN)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)
    normal.font.size = Pt(11)
    normal.font.color.rgb = rgb(INK)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.333

    h_specs = {
        1: (16, MID_BLUE, 18, 10),
        2: (13, MID_BLUE, 12, 6),
        3: (12, BLUE, 8, 4),
    }
    for level, (size, color, before, after) in h_specs.items():
        style = styles[f"Heading {level}"]
        style.font.name = FONT_LATIN
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT_LATIN)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT_LATIN)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = rgb(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    for name in ("List Bullet", "List Bullet 2", "List Number"):
        st = styles[name]
        st.font.name = FONT_LATIN
        st._element.rPr.rFonts.set(qn("w:ascii"), FONT_LATIN)
        st._element.rPr.rFonts.set(qn("w:hAnsi"), FONT_LATIN)
        st._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)
        st.font.size = Pt(10.6)
        st.paragraph_format.space_after = Pt(4)
        st.paragraph_format.line_spacing = 1.208

    header = sec.header
    hp = header.paragraphs[0]
    hp.paragraph_format.space_after = Pt(0)
    rr = hp.add_run("渠道路演材料  |  泛FOF型多资产理财产品")
    set_run_font(rr, size=8.5, bold=True, color=GRAY)
    footer = sec.footer
    add_page_number(footer.paragraphs[0])


def make_charts():
    font_path = "/System/Library/Fonts/STHeiti Light.ttc"
    title_font = ImageFont.truetype(font_path, 31)
    center_font = ImageFont.truetype(font_path, 40)
    center_small = ImageFont.truetype(font_path, 20)
    legend_font = ImageFont.truetype(font_path, 18)

    def donut(values, labels, colors, title, filename, center):
        image = Image.new("RGB", (900, 690), "white")
        draw = ImageDraw.Draw(image)
        draw.text((450, 36), title, fill=f"#{INK}", font=title_font, anchor="mm")
        box = (235, 105, 665, 535)
        start = -90
        for value, color in zip(values, colors):
            end = start + 360 * value / sum(values)
            draw.pieslice(box, start=start, end=end, fill=color, outline="white", width=3)
            start = end
        draw.ellipse((355, 225, 545, 415), fill="white")
        draw.text((450, 296), center[0], fill=f"#{BLUE}", font=center_font, anchor="mm")
        draw.text((450, 344), center[1], fill=f"#{GRAY}", font=center_small, anchor="mm")
        positions = [(85, 565), (465, 565), (85, 610), (465, 610), (85, 655)]
        for idx, (label, color) in enumerate(zip(labels, colors)):
            x, y = positions[idx]
            draw.rounded_rectangle((x, y - 11, x + 22, y + 11), radius=4, fill=color)
            draw.text((x + 32, y), label, fill=f"#{GRAY}", font=legend_font, anchor="lm")
        path = ASSET_DIR / filename
        image.save(path)
        return path

    product = donut(
        [97.09, 2.61, 0.24, 0.06],
        ["固定收益类 97.09%", "混合类 2.61%", "权益类 0.24%", "商品及衍生品类 0.06%"],
        [f"#{BLUE}", f"#{TEAL}", f"#{GOLD}", "#A9B6C2"],
        "理财产品投资性质结构（2025年末）",
        "product_nature.png",
        ("97.09%", "固定收益类"),
    )
    risk = donut(
        [27.84, 67.89, 4.03, 0.12, 0.12],
        ["一级 27.84%", "二级 67.89%", "三级 4.03%", "四级 0.12%", "五级 0.12%"],
        ["#A9C6D8", f"#{BLUE}", f"#{TEAL}", f"#{GOLD}", "#B65C5C"],
        "理财产品风险等级结构（2025年末）",
        "risk_level.png",
        ("95.73%", "二级及以下"),
    )
    return product, risk


def make_example_charts():
    font_path = "/Applications/Microsoft Word.app/Contents/Resources/DFonts/msyh.ttc"
    if not Path(font_path).exists():
        font_path = "/System/Library/Fonts/Hiragino Sans GB.ttc"
    f_title = ImageFont.truetype(font_path, 34)
    f_axis = ImageFont.truetype(font_path, 21)
    f_small = ImageFont.truetype(font_path, 18)
    f_label = ImageFont.truetype(font_path, 23)
    palette = [f"#{BLUE}", f"#{TEAL}", f"#{GOLD}", "#7A8FA6", "#B45B5B", "#6B8E5E"]

    def canvas(title):
        im = Image.new("RGB", (1400, 760), "white")
        d = ImageDraw.Draw(im)
        d.text((700, 44), title, font=f_title, fill=f"#{INK}", anchor="mm")
        return im, d

    def save(im, name):
        p = ASSET_DIR / name
        im.save(p, quality=95)
        return p

    # 1. Strategy clustering map
    im, d = canvas("基金与专户统一聚类示例")
    left, top, right, bottom = 120, 115, 1300, 660
    d.line((left, bottom, right, bottom), fill="#8493A1", width=3)
    d.line((left, bottom, left, top), fill="#8493A1", width=3)
    for i in range(1, 5):
        x = left + (right - left) * i / 5
        y = bottom - (bottom - top) * i / 5
        d.line((x, top, x, bottom), fill="#E3E8ED", width=2)
        d.line((left, y, right, y), fill="#E3E8ED", width=2)
    d.text(((left + right) / 2, 710), "方向性与权益Beta暴露 →", font=f_axis, fill=f"#{GRAY}", anchor="mm")
    d.text((32, (top + bottom) / 2), "波动与回撤风险", font=f_axis, fill=f"#{GRAY}", anchor="mm")
    clusters = [
        ("流动性", 12, 12, 42, 0), ("利率久期", 25, 28, 55, 0),
        ("信用票息", 34, 39, 62, 0), ("绝对收益", 23, 52, 58, 1),
        ("红利低波", 48, 49, 66, 1), ("黄金商品", 55, 61, 60, 2),
        ("可转债", 65, 65, 64, 2), ("科技成长", 82, 84, 72, 4),
    ]
    for name, x0, y0, radius, ci in clusters:
        x = left + (right - left) * x0 / 100
        y = bottom - (bottom - top) * y0 / 100
        d.ellipse((x-radius, y-radius, x+radius, y+radius), fill=palette[ci], outline="white", width=4)
        d.text((x, y), name, font=f_small, fill="white", anchor="mm")
    d.text((1265, 92), "气泡大小：策略容量（示意）", font=f_small, fill=f"#{GRAY}", anchor="ra")
    cluster_map = save(im, "example_cluster_map.png")

    # 2. Typical product portfolio NAVs
    im, d = canvas("典型产品组合净值路径示例")
    left, top, right, bottom = 110, 120, 1320, 650
    d.line((left, bottom, right, bottom), fill="#8493A1", width=3)
    d.line((left, bottom, left, top), fill="#8493A1", width=3)
    for i in range(6):
        y = top + (bottom - top) * i / 5
        d.line((left, y, right, y), fill="#E3E8ED", width=2)
        d.text((95, y), f"{1.16-i*0.04:.2f}", font=f_small, fill=f"#{GRAY}", anchor="rm")
    rng = Random(27)
    specs = [
        ("90天稳健增利", 0.0032, 0.0038, palette[0]),
        ("180天均衡配置", 0.0040, 0.0065, palette[1]),
        ("365天长期增强", 0.0048, 0.0090, palette[2]),
        ("合成指数基准", 0.0035, 0.0046, "#8A929A"),
    ]
    for si, (name, drift, noise, color) in enumerate(specs):
        vals = [1.0]
        for m in range(1, 37):
            shock = -0.018 * max(0, 1 - abs(m-16)/3) - 0.012 * max(0, 1 - abs(m-28)/2)
            seasonal = sin(m * (0.55 + si * 0.04)) * noise
            vals.append(vals[-1] * (1 + drift + seasonal + shock / 4 + rng.uniform(-noise, noise) / 2))
        points = []
        for i, v in enumerate(vals):
            x = left + (right-left) * i / (len(vals)-1)
            y = bottom - (v-0.96) / (1.16-0.96) * (bottom-top)
            points.append((x, y))
        d.line(points, fill=color, width=5 if si < 3 else 3)
        lx = 190 + si * 300
        d.line((lx, 704, lx+55, 704), fill=color, width=5)
        d.text((lx+68, 704), name, font=f_small, fill=f"#{INK}", anchor="lm")
    nav_chart = save(im, "example_typical_nav.png")

    # 3. Product matrix risk-return map
    im, d = canvas("产品矩阵风险收益定位示例")
    left, top, right, bottom = 140, 120, 1290, 650
    d.line((left, bottom, right, bottom), fill="#8493A1", width=3)
    d.line((left, bottom, left, top), fill="#8493A1", width=3)
    for i in range(1, 6):
        x = left + (right-left)*i/6
        y = bottom - (bottom-top)*i/6
        d.line((x, top, x, bottom), fill="#E3E8ED", width=2)
        d.line((left, y, right, y), fill="#E3E8ED", width=2)
    d.text(((left+right)/2, 710), "目标波动率（由低到高） →", font=f_axis, fill=f"#{GRAY}", anchor="mm")
    d.text((48, (top+bottom)/2), "持有期平均收益", font=f_axis, fill=f"#{GRAY}", anchor="mm")
    points = [
        ("低波30天", 14, 20, 34, 0), ("增利90天", 28, 35, 42, 0),
        ("增利180天", 40, 46, 49, 1), ("成长180天", 59, 62, 58, 1),
        ("成长365天", 72, 75, 65, 2), ("增强365天", 86, 86, 75, 4),
    ]
    for name, x0, y0, r0, ci in points:
        x = left+(right-left)*x0/100
        y = bottom-(bottom-top)*y0/100
        d.ellipse((x-r0, y-r0, x+r0, y+r0), fill=palette[ci], outline="white", width=4)
        d.text((x, y), name, font=f_small, fill="white", anchor="mm")
    d.text((1260, 92), "气泡大小：最大回撤（示意）", font=f_small, fill=f"#{GRAY}", anchor="ra")
    matrix_map = save(im, "example_matrix_map.png")

    # 4. High-volatility style blocks
    im, d = canvas("同为10%高波动资产的“积木块”构成示例")
    labels = ["防御收益型", "均衡配置型", "成长增强型"]
    parts = ["红利低波", "宽基Beta", "科技成长", "转债", "黄金商品", "REITs"]
    vals = [[6,1,0,1,1,1], [3,2,2,1,1,1], [0,1,7,1,1,0]]
    left, bar_left, bar_right = 95, 310, 1300
    for row, (label, arr) in enumerate(zip(labels, vals)):
        y = 190 + row*155
        d.text((left, y+34), label, font=f_label, fill=f"#{INK}", anchor="lm")
        x = bar_left
        for j, v in enumerate(arr):
            if v == 0:
                continue
            w = (bar_right-bar_left)*v/10
            d.rectangle((x, y, x+w, y+68), fill=palette[j], outline="white", width=3)
            if w > 70:
                d.text((x+w/2, y+34), f"{parts[j]} {v}%", font=f_small, fill="white", anchor="mm")
            x += w
    for j, part in enumerate(parts):
        x = 150 + (j % 3)*410
        y = 660 + (j // 3)*40
        d.rectangle((x, y-10, x+22, y+12), fill=palette[j])
        d.text((x+32, y), part, font=f_small, fill=f"#{GRAY}", anchor="lm")
    style_blocks = save(im, "example_style_blocks.png")

    # 5. Holding-period positive-return probability
    im, d = canvas("最短持有期与推荐持有期正收益概率示例")
    left, top, right, bottom = 130, 120, 1300, 650
    d.line((left, bottom, right, bottom), fill="#8493A1", width=3)
    d.line((left, bottom, left, top), fill="#8493A1", width=3)
    for pct in range(0, 101, 20):
        y = bottom-(bottom-top)*pct/100
        d.line((left, y, right, y), fill="#E3E8ED", width=2)
        d.text((105, y), f"{pct}%", font=f_small, fill=f"#{GRAY}", anchor="rm")
    cats = ["30天", "90天", "180天", "365天"]
    a, b = [55, 68, 76, 82], [66, 78, 85, 90]
    group = (right-left)/4
    for i, cat in enumerate(cats):
        cx = left+group*(i+0.5)
        for off, val, color in [(-42, a[i], palette[0]), (42, b[i], palette[2])]:
            x1, x2 = cx+off-33, cx+off+33
            y = bottom-(bottom-top)*val/100
            d.rounded_rectangle((x1, y, x2, bottom), radius=7, fill=color)
            d.text(((x1+x2)/2, y-20), f"{val}%", font=f_small, fill=color, anchor="mm")
        d.text((cx, 690), cat, font=f_axis, fill=f"#{INK}", anchor="mm")
    d.rectangle((420, 88, 442, 110), fill=palette[0]); d.text((453, 99), "最短持有期窗口", font=f_small, fill=f"#{GRAY}", anchor="lm")
    d.rectangle((720, 88, 742, 110), fill=palette[2]); d.text((753, 99), "推荐持有期窗口", font=f_small, fill=f"#{GRAY}", anchor="lm")
    holding_chart = save(im, "example_holding_probability.png")

    # 6. Strategy-unit attribution
    im, d = canvas("策略“积木块”收益贡献归因示例")
    items = [("固收票息与久期", 2.20), ("红利低波", 0.38), ("科技成长", 0.16),
             ("黄金商品", 0.24), ("绝对收益", 0.31), ("交易与费用", -0.17)]
    zero = 420
    d.line((zero, 125, zero, 665), fill="#8493A1", width=3)
    for i, (name, val) in enumerate(items):
        y = 160+i*82
        d.text((390, y+25), name, font=f_axis, fill=f"#{INK}", anchor="rm")
        scale = 320
        x2 = zero + val*scale
        color = palette[i % 5] if val >= 0 else "#B45B5B"
        d.rounded_rectangle((min(zero,x2), y, max(zero,x2), y+50), radius=7, fill=color)
        d.text((x2 + (18 if val >= 0 else -18), y+25), f"{val:+.2f}%", font=f_small, fill=color, anchor="lm" if val >= 0 else "rm")
    d.text((1280, 705), "示例区间总收益贡献约3.12%", font=f_axis, fill=f"#{GRAY}", anchor="ra")
    attribution = save(im, "example_attribution.png")

    # 7. Actual NAV vs synthetic benchmark
    im, d = canvas("实际净值与合成指数基准跟踪示例")
    left, top, right, bottom = 110, 130, 1320, 650
    d.line((left, bottom, right, bottom), fill="#8493A1", width=3)
    d.line((left, bottom, left, top), fill="#8493A1", width=3)
    for i in range(5):
        y = top+(bottom-top)*i/4
        d.line((left, y, right, y), fill="#E3E8ED", width=2)
    benchmark, actual = [1.0], [1.0]
    for m in range(1, 31):
        base = 0.0035 + sin(m*0.6)*0.0025 - 0.010*max(0, 1-abs(m-17)/2)
        benchmark.append(benchmark[-1]*(1+base))
        alpha = 0.0007 - (0.0014 if 12 <= m <= 16 else 0)
        actual.append(actual[-1]*(1+base+alpha))
    for vals, color, width in [(benchmark, "#7A8FA6", 4), (actual, palette[0], 6)]:
        pts=[]
        for i,v in enumerate(vals):
            x=left+(right-left)*i/(len(vals)-1)
            y=bottom-(v-0.98)/(1.14-0.98)*(bottom-top)
            pts.append((x,y))
        d.line(pts, fill=color, width=width)
    d.line((430, 95, 490, 95), fill=palette[0], width=6); d.text((505,95),"实际净值",font=f_small,fill=f"#{INK}",anchor="lm")
    d.line((735, 95, 795, 95), fill="#7A8FA6", width=4); d.text((810,95),"合成指数基准",font=f_small,fill=f"#{INK}",anchor="lm")
    d.rectangle((left+(right-left)*12/30, top, left+(right-left)*16/30, bottom), outline=f"#{GOLD}", width=3)
    d.text((left+(right-left)*14/30, top+30), "阶段性偏离诊断区", font=f_small, fill=f"#{GOLD}", anchor="mm")
    tracking = save(im, "example_benchmark_tracking.png")

    return {
        "cluster": cluster_map,
        "nav": nav_chart,
        "matrix": matrix_map,
        "style": style_blocks,
        "holding": holding_chart,
        "attribution": attribution,
        "tracking": tracking,
    }


def add_picture_pair(doc, left, right):
    table = doc.add_table(rows=1, cols=2)
    set_table_geometry(table, [4680, 4680], indent=0)
    for cell, path in zip(table.rows[0].cells, [left, right]):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.add_run().add_picture(str(path), width=Inches(3.0))
        set_cell_margins(cell, top=20, bottom=20, start=40, end=40)
    # no visible layout borders
    set_table_borders(table, color=WHITE, size=0)


def add_metric_strip(doc, metrics):
    table = doc.add_table(rows=1, cols=len(metrics))
    widths = [9360 // len(metrics)] * len(metrics)
    widths[-1] += 9360 - sum(widths)
    set_table_geometry(table, widths, indent=0)
    set_table_borders(table, color=WHITE, size=0)
    for cell, (value, label) in zip(table.rows[0].cells, metrics):
        set_cell_shading(cell, PALE_BLUE)
        set_cell_margins(cell, top=150, bottom=150, start=100, end=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(value)
        set_run_font(r, size=17, bold=True, color=BLUE)
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(label)
        set_run_font(r2, size=8.8, color=GRAY)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def page_break(doc):
    doc.add_page_break()


def build():
    product_chart, risk_chart = make_charts()
    examples = make_example_charts()
    doc = Document()
    configure_styles(doc)

    # Cover
    add_para(doc, "渠道合作路演材料", size=11, bold=True, color=GOLD,
             align=WD_ALIGN_PARAGRAPH.LEFT, before=30, after=26)
    add_para(doc, "泛FOF型多资产\n理财产品方案", size=29, bold=True, color=BLUE,
             align=WD_ALIGN_PARAGRAPH.LEFT, before=40, after=14, line=1.05)
    add_para(doc, "用专业“积木块”搭建持有期体验导向的产品矩阵", size=15, color=TEAL,
             align=WD_ALIGN_PARAGRAPH.LEFT, after=24)
    add_callout(
        doc,
        "面向银行财富渠道的核心价值",
        "用“风险风格 × 最短持有期”形成清晰货架；将基金、专户和策略统一为可识别、可组合、可归因的专业“积木块”；用正收益概率、收益分布、回撤和修复周期管理客户持有体验。",
        fill=PALE_BLUE,
        accent=TEAL,
    )
    add_para(doc, "适用读者：渠道产品经理、销售管理人员、财富顾问团队", size=10.5,
             color=GRAY, align=WD_ALIGN_PARAGRAPH.LEFT, before=40, after=5)
    add_para(doc, "版本：V1.1  |  2026年7月", size=10, color=GRAY,
             align=WD_ALIGN_PARAGRAPH.LEFT, after=3)
    add_para(doc, "本材料用于产品方案交流，不构成任何收益承诺或投资建议。", size=9,
             color=RED, align=WD_ALIGN_PARAGRAPH.LEFT, before=35, after=0)

    page_break(doc)
    add_heading(doc, "一、执行摘要：渠道需要一套“看得懂、卖得准、拿得住”的产品体系", 1)
    add_metric_strip(doc, [
        ("33.29万亿元", "2025年末理财存续规模"),
        ("1.43亿个", "持有理财产品的投资者数量"),
        ("97.09%", "固定收益类产品规模占比"),
        ("95.73%", "二级及以下产品规模占比"),
    ])
    add_source(doc, "《中国银行业理财市场年度报告（2025年）》；数据截至2025年末。")
    add_para(doc, "银行理财市场规模持续扩大，但产品结构仍高度集中于固定收益类和中低风险产品。渠道真正缺少的，不是更多名称相似的“固收+”，而是一套能够清楚解释风险来源、持有期限和客户体验差异的产品矩阵。")
    add_callout(
        doc,
        "方案结论",
        "以固收底仓形成基础收益和流动性缓冲，以多资产、多策略和多风格“积木块”获取风险溢价；先确定客户风险与持有期，再配置资产簇和策略簇，最后选择公募基金、专户及具体标的。",
    )
    add_bullets(doc, [
        "四类风险风格：稳健低波、稳健增利、均衡成长、进取增强。",
        "四档最短持有期：30天、90天、180天、365天。",
        "每个矩阵单元是一套独立目标组合，而不是同一组合更换期限标签。",
        "同时设置最短持有期和推荐持有期，分别管理赎回安排与适宜投资周期。",
        "统一对公募基金和专户进行资产、策略、风格与风险因子分类，形成可替换、可评价的“积木块”库。",
    ])

    page_break(doc)
    add_heading(doc, "二、市场背景：大体量、低风险偏好与多资产转型并存", 1)
    add_picture_pair(doc, product_chart, risk_chart)
    add_source(doc, "银行业理财登记托管中心，《中国银行业理财市场年度报告（2025年）》。")
    add_para(doc, "2025年末，固定收益类产品存续规模32.32万亿元，占比97.09%；风险等级二级及以下产品规模31.87万亿元，占比95.73%。个人投资者1.41亿个，占全部理财投资者的98.64%，其中稳健型投资者数量占比最高，达到33.54%。")
    add_para(doc, "与此同时，行业正在从“单一资产驱动”向“多资产配置”转型。年度报告中的行业案例显示，头部机构已开始使用宏观因子、风险预算、情景分析、策略组合和工具层对冲，推动多资产配置从简单拼接走向组合管理。")
    add_callout(
        doc,
        "渠道机会",
        "客户风险偏好整体稳健，并不意味着产品只能依赖单一固收收益。更可持续的方向，是在严格风险预算下引入差异化资产和策略，通过分散风险来源改善持有期结果。",
        fill=PALE_GOLD,
        accent=GOLD,
    )

    page_break(doc)
    add_heading(doc, "三、公开案例：持有期产品已形成清晰的期限梯度", 1)
    add_para(doc, "公开产品文件显示，最短持有期已经成为开放式理财产品常见的流动性安排。市场上已存在30天、90天、180天和365天等不同期限产品，为建立期限矩阵提供了成熟的客户认知基础。")
    add_table(
        doc,
        ["期限案例", "公开产品案例", "可借鉴要点"],
        [
            ["30天", "信银理财安盈象固收稳利一个月持有期101号", "每笔份额持有满30个自然日后可在开放日申请赎回"],
            ["90天", "中银理财－悦享（90天持有期）", "最短持有90个自然日，产品成立后工作日开放申购赎回"],
            ["180天", "工银理财·核心优选最短持有180天固定收益类开放式产品", "公募、开放式、固定收益类、PR2；基准由存款与同业存单指数等合成"],
            ["365天", "中邮理财灵活添利·鸿锦最短持有365天8号", "以365天持有期承载更长的策略实现周期"],
        ],
        [1100, 4700, 3560],
        font_size=8.8,
        first_col_bold=True,
    )
    add_source(doc, "招商银行代销产品页面、中银理财产品说明书、工银理财发行公告、中邮理财发行公告；详见文末资料来源。")
    add_para(doc, "本方案不把“最短持有期”等同于“推荐持有期”。客户满期后获得赎回权，但是否适合继续持有，应由组合波动、收益分布和回撤修复周期决定。")
    add_placeholder(
        doc, "T-01", "期限梯度与客户持有体验",
        "形成30/90/180/365天产品在渠道端的直观对比。",
        "使用拟定SAA和代表性指数，分别测算最短持有期与推荐持有期的滚动收益。",
        "期限—正收益概率曲线、期限—5%分位收益曲线、期限—回撤修复概率曲线。",
    )

    page_break(doc)
    add_heading(doc, "四、产品定位：不是公募FOF的复制，而是理财特色的泛FOF", 1)
    add_para(doc, "泛FOF是一种组合管理方法：在产品层配置多资产、多策略、多风格和多管理人，并通过公募基金、专户、其他合规资产管理计划及允许范围内的直接投资加以实现。")
    add_table(
        doc,
        ["比较维度", "泛FOF型理财产品", "公募FOF"],
        [
            ["风险偏好", "整体偏中低风险，强调固收底仓、持有期正收益概率和回撤控制", "风险谱系更宽，可覆盖偏债、平衡、偏股、目标风险及养老目标等类型"],
            ["产品起点", "银行客户风险偏好、资金期限和持有体验", "基金合同约定的投资目标和风险收益特征"],
            ["底层工具", "多资产、多策略公募基金、专户、资管计划及合规直投", "原则上80%以上基金资产投资于其他公募基金份额"],
            ["固收实施", "可直接管理固收资产，亦可通过定制专户和公募工具实施", "主要通过债券基金、货币基金等间接实现"],
            ["策略定制", "可直接设定久期、信用、杠杆、风格、流动性和风险边界", "受底层基金合同、开放安排与管理方式约束"],
            ["客户管理", "最短持有期＋推荐持有期＋银行渠道陪伴", "依据基金合同、销售适当性和基金持有安排"],
        ],
        [1500, 3930, 3930],
        font_size=8.4,
        first_col_bold=True,
    )
    add_source(doc, "中国证监会《公开募集证券投资基金运作指引第2号——基金中基金指引》；银行理财相关制度。")
    add_callout(
        doc,
        "差异化价值",
        "本方案的优势不是简单增加权益仓位，而是在较低风险预算下，使用更丰富的底层工具和策略组合，争取更稳定、更可解释的持有期体验。",
    )

    page_break(doc)
    add_heading(doc, "五、产品矩阵：风险风格 × 最短持有期", 1)
    add_table(
        doc,
        ["风险风格", "30天", "90天", "180天", "365天"],
        [
            ["稳健低波（高波动资产0%–5%）", "核心", "扩展", "扩展", "储备"],
            ["稳健增利（高波动资产0%–10%）", "储备", "核心", "核心", "扩展"],
            ["均衡成长（高波动资产5%–15%）", "不建议", "储备", "核心", "核心"],
            ["进取增强（高波动资产5%–25%）", "不建议", "不建议", "储备", "核心"],
        ],
        [3600, 1440, 1440, 1440, 1440],
        font_size=8.7,
        first_col_bold=True,
    )
    add_para(doc, "“核心”代表首期或重点布局；“扩展”代表具备渠道需求后逐步发行；“储备”代表保留策略模型；“不建议”代表风险特征与最短持有期明显不匹配。")
    add_callout(
        doc,
        "每个矩阵单元的定义",
        "产品矩阵单元 = 资产簇配置 × 策略与风格“积木块” × 实施工具 × 风险预算 × 流动性约束 × 持有期目标。",
        fill=PALE_GOLD,
        accent=GOLD,
    )
    add_figure(
        doc,
        examples["matrix"],
        "示例图：六类核心产品在目标波动、持有期收益和最大回撤三个维度的相对定位",
        "模拟数据，仅用于展示图表结构与阅读方式，不代表拟发行产品的收益预测。",
        width=6.15,
    )
    add_para(doc, "示例阅读：从左下到右上依次对应更高的风险预算和更长的策略实现周期；同一期限下，产品仍可因风格“积木块”、固收久期和对冲安排不同而处于不同位置。", size=9.3, color=GRAY)
    add_placeholder(
        doc, "T-02", "完整产品矩阵风险收益地图",
        "让渠道直观看到各矩阵单元的收益目标、波动与回撤差异。",
        "基于各单元SAA、费用和历史代表指数测算，统一采用最短持有期与推荐持有期两个窗口。",
        "二维气泡图：横轴为目标波动，纵轴为持有期平均收益，气泡大小代表最大回撤或尾部损失。",
    )

    add_heading(doc, "六、首期核心产品：形成清晰可销售的六个组合单元", 1)
    add_table(
        doc,
        ["产品单元", "最短持有期", "推荐持有期", "高波动资产", "战略中枢", "核心定位"],
        [
            ["稳健低波30天", "30天", "3个月以上", "0%–5%", "2%", "短持有、低波动、流动性优先"],
            ["稳健增利90天", "90天", "6个月以上", "0%–10%", "5%", "固收底仓加有限多策略增强"],
            ["稳健增利180天", "180天", "9–12个月", "0%–10%", "6%", "增加固收策略空间与收益弹性"],
            ["均衡成长180天", "180天", "12个月以上", "5%–15%", "10%", "多风格权益及其他风险溢价组合"],
            ["均衡成长365天", "365天", "18个月以上", "5%–15%", "12%", "更完整的跨资产、跨周期配置"],
            ["进取增强365天", "365天", "24个月以上", "5%–25%", "15%", "长期风险溢价与主动管理结合"],
        ],
        [1940, 1120, 1350, 1400, 1050, 2500],
        font_size=8.15,
        first_col_bold=True,
    )
    add_para(doc, "战略中枢代表正常市场状态下的目标位置，高波动资产上下限则为战术调整和风险控制提供空间。最终中枢、推荐持有期及监管分类须以回测、压力测试和正式立项结果为准。")
    add_placeholder(
        doc, "T-03", "六类产品持有期画像表",
        "形成渠道可直接使用的产品比较页。",
        "对每个产品分别统计最短持有期和推荐持有期的滚动结果，并扣除拟定综合费率。",
        "正收益概率、平均收益、中位数、5%/25%/75%/95%分位、单位净值波动率、最大回撤、修复时间。",
    )
    add_heading(doc, "典型产品案例库：沉淀可复用的“积木搭法”", 2)
    add_table(
        doc,
        ["案例", "典型客户场景", "核心“积木块”", "组合与管理重点"],
        [
            ["案例A｜稳健低波30天", "短期闲置资金，重视流动性和净值稳定", "短久期高等级固收专户＋货币工具＋少量绝对收益", "控制久期、杠杆和解锁日前流动性；高波动资产以0%–2%为中枢"],
            ["案例B｜稳健增利90天", "可承受轻微波动，希望高于纯现金管理体验", "中短久期信用专户＋红利低波＋黄金", "固收贡献主要收益；估值与拥挤度决定红利预算，黄金用于尾部分散"],
            ["案例C｜稳健增利180天", "资金半年以上不用，重视正收益概率和回撤修复", "票息与骑乘专户＋转债＋红利＋市场中性", "通过收益积累释放风险预算；转债估值过高时由绝对收益策略替代"],
            ["案例D｜均衡成长180天", "接受阶段性浮亏，希望兼顾防御和上涨弹性", "固收底仓＋红利低波＋宽基指数增强＋科技成长＋黄金", "高波动资产约10%；控制成长风格集中度及股债相关性上升风险"],
            ["案例E｜均衡成长365天", "一年以上配置资金，关注多资产跨周期能力", "固收专户＋多风格权益＋转债＋REITs＋商品", "允许更完整的风险预算配置；通过估值、趋势和流动性信号进行再平衡"],
            ["案例F｜进取增强365天", "较高风险承受能力，接受较大净值波动", "指数增强＋主动成长＋量化选股＋黄金商品＋对冲策略", "强化风格和容量管理；设置风险预算上限、压力情景和分阶段恢复条件"],
        ],
        [1800, 2400, 2800, 2360],
        font_size=7.8,
        first_col_bold=True,
    )
    add_para(doc, "案例库用于沉淀可复用的产品设计逻辑：每个案例均保留“客户场景—积木块构成—模型参数—组合结果—销售表达—投后复盘”六类档案。正式产品可在同一案例框架下进行簇内标的替换，而不改变产品定位。", size=9.3, color=GRAY)

    page_break(doc)
    add_heading(doc, "七、“积木块”体系：统一分类基金与专户，再搭建产品组合", 1)
    add_callout(
        doc,
        "什么是“积木块”",
        "“积木块”是完成资产归属、策略标签、风格画像、风险因子、流动性约束和准入评价后的可组合投资单元。它既可以是一只公募基金，也可以是一个定制专户或其他合规资管计划；载体不同，但必须使用同一套语言识别其真实风险来源。",
        fill=PALE_GOLD,
        accent=GOLD,
    )
    add_table(
        doc,
        ["资产簇", "主要资产与策略", "组合功能"],
        [
            ["流动性资产簇", "现金、存款、同业存单、货币工具、短期回购", "申赎准备、净值稳定、交易缓冲"],
            ["利率资产簇", "国债、政策性金融债、利率债基金及专户", "久期收益、流动性、避险"],
            ["信用资产簇", "高等级信用债、信用债基金及专户", "票息和信用利差收益"],
            ["固收增强簇", "骑乘、杠杆、信用挖掘、波段", "提高固收收益弹性"],
            ["权益Beta/Alpha簇", "指数、指数增强、主动权益、量化选股", "获取市场风险溢价与管理人Alpha"],
            ["转债与实物资产簇", "可转债、黄金、商品、公募REITs", "增强、通胀对冲和风险分散"],
            ["绝对收益与对冲簇", "市场中性、多策略、套利及合规衍生品", "降低方向依赖、管理尾部风险"],
        ],
        [1800, 4300, 3260],
        font_size=8.6,
        first_col_bold=True,
    )
    add_para(doc, "公募基金和专户首先按照同一套资产、策略和风格标签分类，再比较费率、透明度、流动性、合同约束和管理人能力。载体不同，不应被误认为风险来源已经分散。")
    add_heading(doc, "从全市场标的到代表性“积木块”", 2)
    add_numbered(doc, [
        "统一分类：按资产、策略、风格和风险因子完成基金与专户的共同标签。",
        "策略聚类：以收益相关性、波动、回撤、Beta、风格暴露和持仓特征识别同类策略，减少“名称不同、风险相同”的重复配置。",
        "簇内精选：在同一聚类内比较持续Alpha、下行保护、费率、流动性、透明度、容量和管理稳定性。",
        "组合搭建：从不同簇中选择互补“积木块”，形成典型产品组合并回测净值与持有期结果。",
        "动态替换：当管理人、风格或执行效率变化时，在不改变产品定义的前提下优先进行簇内替换。",
    ])
    add_callout(
        doc,
        "组合纪律",
        "先回答组合需要什么风险收益来源，再回答用哪只基金、哪个专户和哪位管理人实现。优秀的单一标的，只有在改善组合持有期结果时才具有入池价值。",
    )
    add_figure(
        doc,
        examples["cluster"],
        "示例图：基金与专户采用统一风险收益特征进行聚类",
        "模拟数据，仅展示聚类后的策略簇、相对风险和容量表达方式。",
        width=6.15,
    )
    add_para(doc, "示例阅读：同一聚类内的基金和专户承担相似风险来源，可进行费率、Alpha、流动性和管理稳定性的簇内比较；组合分散应优先跨簇实现，而不是机械增加同簇标的数量。", size=9.3, color=GRAY)
    add_asset_slot(
        doc, "A-01", "统一分类、策略聚类与代表性“积木块”地图",
        "置于本节末，用一张图展示从标的池、聚类到代表性策略单元的筛选路径。",
        "基金与专户使用相同观察期、收益频率、费后净值和风险因子定义；聚类算法、距离度量及人工复核规则留痕。",
        "聚类树或二维降维散点图＋各簇标签、代表标的、簇内离散度和容量。",
    )
    add_figure(
        doc,
        examples["nav"],
        "示例图：典型产品组合与合成指数基准的净值路径",
        "模拟数据，统一起点为1；仅用于展示正式组合回测图的结构。",
        width=6.15,
    )
    add_para(doc, "示例阅读：比较期末收益之外，还应同步观察回撤发生时间、各组合下跌幅度、恢复速度及与合成基准的阶段性偏离。", size=9.3, color=GRAY)
    add_asset_slot(
        doc, "A-02", "典型产品组合构成与回测净值",
        "紧接聚类地图，展示代表性“积木块”如何搭成六类核心产品。",
        "统一起始日、净值归一为1、费率假设、再平衡频率、现金拖累和可交易性；明确实盘、回测及代理指数区间。",
        "组合权重堆叠图＋组合净值、合成基准与回撤子图＋最短/推荐持有期统计摘要。",
    )

    page_break(doc)
    add_heading(doc, "八、同样10%的高波动资产，不同“积木块”会形成不同产品特征", 1)
    add_table(
        doc,
        ["策略子簇", "防御收益型", "均衡配置型", "成长增强型"],
        [
            ["红利价值", "4%", "2%", "0%"],
            ["低波动权益", "2%", "1%", "0%"],
            ["宽基Beta", "1%", "2%", "1%"],
            ["科技成长", "0%", "2%", "5%"],
            ["小盘成长", "0%", "0%", "2%"],
            ["转债", "1%", "1%", "1%"],
            ["黄金及商品", "1%", "1%", "1%"],
            ["公募REITs", "1%", "1%", "0%"],
            ["合计", "10%", "10%", "10%"],
        ],
        [3000, 2120, 2120, 2120],
        font_size=8.7,
        first_col_bold=True,
    )
    add_para(doc, "三种方案名义仓位相同，但风险来源不同：防御收益型更依赖股息、低波和实物资产分散；成长增强型更依赖科技、小盘和高Beta环境；均衡配置型则在多种风格之间分配风险预算。")
    add_figure(
        doc,
        examples["style"],
        "示例图：同为10%高波动资产，不同风格“积木块”的组合方式",
        "根据本节示例权重绘制；不构成正式产品配置建议。",
        width=6.15,
    )
    add_placeholder(
        doc, "T-04", "高波动资产风格指纹",
        "展示产品并非只控制权益比例，还控制风格、行业和风险因子。",
        "对底层公募基金和专户进行穿透，汇总价值/成长、大小盘、红利、质量、动量、Beta、行业及流动性暴露。",
        "六个核心产品的风格雷达图＋行业集中度热力图。",
    )

    add_heading(doc, "九、固收底仓与专户：稳定收益来源，也是产品差异化基础", 1)
    add_para(doc, "固收仓承担基础收益积累、净值稳定、流动性管理、回撤缓冲和战术资金来源。不同期限产品应使用不同的久期、信用、杠杆和流动性安排。")
    add_table(
        doc,
        ["产品单元", "固收底仓建议特征"],
        [
            ["稳健低波30天", "短久期、高等级、高流动性；严格限制信用下沉和杠杆"],
            ["稳健增利90天", "中短久期、适度信用挖掘，以票息积累与流动性为主"],
            ["稳健增利180天", "可适度增加骑乘、杠杆、信用利差和资本利得策略"],
            ["均衡成长180天", "承担净值稳定、流动性和高波动仓缓冲"],
            ["均衡成长365天", "可使用更完整的久期、曲线和跨策略组合"],
            ["进取增强365天", "兼顾收益积累、战术资金和组合尾部缓冲"],
        ],
        [2300, 7060],
        font_size=9,
        first_col_bold=True,
    )
    add_callout(
        doc,
        "专户的价值",
        "通过合同定制久期、信用、杠杆、流动性、预警和报告要求；在规模允许时争取更具竞争力的费率和更透明的风险管理。专户是否优于公募基金，必须综合管理费、托管费、运营成本、最低规模和流动性成本比较。",
    )
    add_heading(doc, "固收专户相对指数工具的实施案例", 2)
    add_para(doc, "以同区间、同起点净值比较固收专户、债券指数和债券基金指数，可以更直观地检验专户在收益获取、回撤控制和路径稳定性上的实施效果。评价结论应同时结合策略约束与费用，避免仅凭期末收益判断。")
    add_figure(
        doc,
        SOURCE_MEDIA / "image12.png",
        "图：固收专户净值与债券指数、债券基金指数对比示例",
        "已有内部策略材料；对外使用前应核对专户名称、样本区间、基准名称、费后口径及数据授权。",
        width=6.35,
    )
    add_bullets(doc, [
        "收益维度：区间收益、年化收益、超额收益及滚动胜率。",
        "风险维度：年化波动率、最大回撤、回撤修复时间、下行捕获率和尾部损失。",
        "实施维度：综合费率、久期与信用偏离、杠杆使用、流动性、容量及信息透明度。",
    ])
    add_placeholder(
        doc, "T-05", "公募基金与专户实施效率比较",
        "为各资产簇选择更合适的实施载体。",
        "在相同策略口径下比较公募基金与专户的费率、跟踪误差、Alpha、最大回撤、流动性、透明度和运营成本。",
        "分策略评分卡＋三年情景成本测算＋最小经济规模。",
    )

    page_break(doc)
    add_heading(doc, "十、投资管理：SAA确定产品性格，TAA管理阶段性风险", 1)
    add_numbered(doc, [
        "以客户需求、风险承受能力和持有期限确定产品矩阵单元。",
        "设置收益目标、波动与回撤容忍度、流动性约束和业绩比较基准。",
        "通过SAA确定资产簇、策略簇、风格子簇和长期风险预算。",
        "通过TAA在授权区间内，根据增长、通胀、流动性、估值、趋势和风险偏好调整。",
        "选择相应策略的公募基金、专户和具体标的，完成穿透式风险复核。",
        "持续监控风险贡献、持有期结果、流动性和客户行为，并按规则再平衡。",
    ])
    add_table(
        doc,
        ["管理层级", "回答的问题", "主要输出"],
        [
            ["SAA战略配置", "产品长期依靠哪些风险收益来源？", "资产中枢、风格预算、风险贡献、长期基准"],
            ["TAA战术配置", "当前市场状态下应偏离多少？", "资产偏离、风险状态、加减仓和恢复条件"],
            ["标的选择", "用什么工具和管理人实现？", "公募基金、专户、指数化工具及具体标的"],
            ["风险与质量控制", "实际组合是否仍符合产品定义？", "限额、预警、压力测试、回撤与流动性处置"],
        ],
        [1800, 3700, 3860],
        font_size=8.9,
        first_col_bold=True,
    )
    add_heading(doc, "风险预算模型示例：量化形成底稿，主观校准关键参数", 2)
    add_para(doc, "风险预算模型将组合总风险分配到不同资产簇和策略“积木块”，根据波动率与相关性求得初始权重。模型输出不是自动交易指令，而是可复核的配置底稿；投研团队结合国内市场结构和当前环境，对波动率、相关性、预期回报、置信度、容量与上下限进行主观校准。")
    add_callout(
        doc,
        "风险预算模型的核心公式",
        "组合波动率：σp = √(wᵀΣw)\n边际风险贡献：MRC_i = (Σw)_i / σp\n策略风险贡献：RC_i = w_i × MRC_i\n相对风险贡献：RRC_i = RC_i / σp\n优化目标：min Σ_i (RRC_i − b_i)²，并满足资金权重和为1、资产上下限、流动性及监管比例等约束。",
        fill=PALE_BLUE,
        accent=TEAL,
    )
    add_para(doc, "其中，w为各“积木块”的资金权重向量，Σ为协方差矩阵，σp为组合波动率，MRC_i为第i个策略的边际风险贡献，RC_i为其绝对风险贡献，RRC_i为相对风险贡献，b_i为经审批的目标风险预算。风险预算模型关注“谁贡献了多少风险”，因此资金权重与风险预算通常并不相等。", size=9.5)
    add_table(
        doc,
        ["“积木块”风险来源", "量化初始风险预算", "主观校准后", "校准考虑（示例）"],
        [
            ["固收票息与久期", "70%", "76%", "流动性宽松、票息可见度较高；同时约束久期和信用尾部"],
            ["红利与低波权益", "10%", "8%", "估值仍有支撑，但交易拥挤度上升"],
            ["科技成长权益", "8%", "4%", "盈利弹性较高，但波动率和风格拥挤度偏高"],
            ["黄金及商品", "7%", "7%", "保留对冲通胀、地缘与股债同跌风险的预算"],
            ["绝对收益与其他策略", "5%", "5%", "结合策略容量、流动性和实盘稳定性控制上限"],
        ],
        [2700, 1700, 1500, 3460],
        font_size=8.25,
        first_col_bold=True,
    )
    add_para(doc, "注：表中为机制示例，风险预算不等同于资金权重，也不代表拟发行产品的最终参数。模型需在产品约束下反解资金权重，并通过压力测试、交易成本和流动性检验。", size=8.8, color=GRAY)
    add_heading(doc, "关键参数如何被“量化＋主观”共同调整", 3)
    add_table(
        doc,
        ["参数", "量化基础值（示例）", "主观校准（示例）", "对组合的影响"],
        [
            ["波动率估计σ_i", "过去3年日收益，60日半衰期EWMA", "进入冲击状态后改为20日半衰期，并设置历史压力下限", "更快反映近期风险，减少模型因长期均值而低估仓位风险"],
            ["相关性矩阵Σ", "历史样本＋收缩估计", "Σ*＝(1−λ)Σ历史＋λΣ压力；λ由20%提高至35%", "提高股债同跌或风格共振情景权重，降低虚假分散"],
            ["目标风险预算b_i", "依据产品SAA和长期风险溢价设定", "结合估值、政策、盈利趋势和拥挤度乘以0.5–1.3的主观系数", "改变风险在红利、成长、商品和绝对收益策略间的分配"],
            ["流动性折扣L_i", "日均成交、赎回安排、冲击成本和容量评分", "小微盘、低流动性信用及容量接近上限时由1.00下调至0.70", "降低可配置上限，增加现金和高流动性工具"],
            ["置信度C_i", "信号稳定性、样本外胜率和模型一致性", "政策或制度变化导致历史规律失真时下调置信度", "缩小TAA偏离幅度，避免过度依赖单一信号"],
            ["仓位上下限", "产品合同、监管分类和长期SAA区间", "根据回撤预算、策略容量及渠道解锁节奏进一步收紧", "保证模型解可执行，并控制客户持有期体验"],
        ],
        [1750, 2450, 2750, 2410],
        font_size=7.7,
        first_col_bold=True,
    )
    add_callout(
        doc,
        "参数合成示例",
        "经调整的协方差矩阵采用Σ*；经调整的目标风险预算可写为 b_i* ∝ b_i量化 × S_i主观 × L_i流动性 × C_i置信度，并重新归一化使Σb_i*=100%。例如，科技成长的量化风险预算为8%，若主观系数0.70、流动性系数0.90、置信度0.85，则调整前乘积为4.28%，再与其他策略共同归一化并接受仓位上下限检验。",
        fill=PALE_GOLD,
        accent=GOLD,
    )
    add_callout(
        doc,
        "“主观＋量化”的决策闭环",
        "量化模型负责一致地测量风险并生成初始方案；主观研究负责识别模型尚未充分反映的制度、估值、政策、拥挤、流动性和尾部变化；风险纪律负责限定调整幅度、记录参数版本，并设定触发、退出与恢复条件。",
        fill=PALE_GOLD,
        accent=GOLD,
    )
    add_heading(doc, "历史市场情景下的参数校准与组合动作示例", 2)
    add_table(
        doc,
        ["历史情景", "市场特征", "量化模型信号", "主观校准与组合动作"],
        [
            ["2018年A股风险偏好下行", "上证综指全年下跌24.59%，深证成指下跌34.42%；股票市场压力处于较高水平", "权益波动率和下行相关性上升，成长与小盘风险贡献快速提高", "提高权益压力波动率，压降成长预算；增加高等级利率债、红利低波和现金缓冲，待估值与趋势共同修复后分阶段恢复"],
            ["2020年初疫情冲击", "全球经济和金融市场遭受突发冲击，跨资产相关性与流动性短期失真", "历史相关性低估共振风险，短窗口波动和流动性指标迅速恶化", "提高压力矩阵权重λ和现金需求假设；优先降低低流动性及高Beta策略，保留高等级固收，并以政策、流动性和趋势信号作为恢复条件"],
            ["2022年11月债市调整", "债券收益率上升、部分理财净值下跌并引发集中赎回，赎回又增加债市调整压力", "固收波动率、久期风险和流动性折价同步上升，原有“低波底仓”风险贡献超预算", "上调固收压力波动率和赎回情景，降低杠杆与久期，增加可变现资产；避免机械抛售，按解锁日和客户行为安排流动性"],
            ["2024年初小微盘冲击", "一季度中证1000、中证2000及微盘风格明显弱于大盘与红利，风格分化扩大", "小盘Beta、拥挤度和流动性因子触发预警，簇内相关性明显升高", "将小微盘流动性折扣由1.00下调至0.70，收紧容量和单一管理人上限；转向宽基、红利及基本面更稳定的指数增强“积木块”"],
            ["2024年9月后风险偏好切换", "9月24日一揽子金融政策提升市场信心和流动性，权益市场快速走强并发生风格切换", "趋势和风险偏好信号转正，但短期波动、相关性和拥挤度同时上升", "不直接按趋势满仓；提高权益预算但分批执行，保留风格分散与止盈再平衡，设置成交量、估值和波动率作为继续加仓或降温条件"],
        ],
        [1550, 2500, 2400, 2910],
        font_size=7.35,
        first_col_bold=True,
    )
    add_para(doc, "上述历史情景用于说明参数校准方法，不代表对未来市场路径的判断。正式回放应使用可获得的当时数据，避免引入事后信息，并记录模型信号产生日、投资委员会决策日和实际交易日。", size=8.8, color=GRAY)
    add_placeholder(
        doc, "T-06", "市场状态与TAA决策面板",
        "将动态管理转化为可解释、可复核的规则。",
        "按增长、通胀、流动性、估值、趋势和波动构建信号，设置正常/预警/防御/恢复四种状态。",
        "状态仪表盘、资产倾向评分、允许偏离范围及历史触发回放。",
    )

    page_break(doc)
    add_heading(doc, "十一、持有期管理：从“能赎回”走向“更适合持有多久”", 1)
    add_table(
        doc,
        ["概念", "客户含义", "管理含义"],
        [
            ["最短持有期", "每笔份额未满期限不能赎回，满期后可按规则申请赎回", "稳定产品规模、支持策略实施和流动性管理"],
            ["推荐持有期", "基于组合特征建议的更适宜观察周期，不构成收益承诺", "用概率分布和回撤修复判断客户适配性"],
        ],
        [1900, 3900, 3560],
        font_size=9,
        first_col_bold=True,
    )
    add_para(doc, "每个产品同时测算两个窗口：滚动持有满最短持有期，以及滚动持有满推荐持有期。渠道看到的不应只有某一时点年化收益，而应包括不同买入时点的完整结果分布。")
    add_bullets(doc, [
        "结果指标：正收益概率、平均收益、中位数收益、5%/25%/75%/95%分位和最差持有期收益。",
        "过程指标：单位净值波动率、窗口内最大回撤、回撤持续时间、修复时间和负收益天数占比。",
        "方法组合：历史滚动窗口、历史情景重演、Bootstrap、Monte Carlo、压力测试和参数敏感性分析。",
    ])
    add_figure(
        doc,
        examples["holding"],
        "示例图：最短持有期窗口与推荐持有期窗口的正收益概率比较",
        "模拟数据，仅用于展示双窗口测算的表达方式；正式结果须基于组合回测与费用口径。",
        width=6.05,
    )
    add_para(doc, "示例阅读：推荐持有期通常比最短持有期更长，目的在于给策略和市场修复留出时间。概率提高不等于承诺盈利，仍需同时展示尾部收益和最大回撤分布。", size=9.3, color=GRAY)
    add_placeholder(
        doc, "T-07", "最短持有期与推荐持有期双窗口测算",
        "支撑产品命名、渠道适配和客户陪伴。",
        "基于正式SAA、费用、再平衡和流动性假设，对所有申购日进行滚动持有测算。",
        "双窗口正收益概率、收益分布箱线图、最大回撤分布、回撤修复概率。",
    )

    add_heading(doc, "十二、全流程质量控制：让产品定义在投资过程中保持一致", 1)
    add_table(
        doc,
        ["环节", "主要质量控制"],
        [
            ["产品立项", "客户需求、产品矩阵、期限和风险目标匹配"],
            ["模型验证", "数据、参数、样本、压力情景和稳健性复核"],
            ["标的准入", "统一策略分类、管理人能力、合同与运营风险审查"],
            ["组合构建", "SAA、风格预算、风险贡献与流动性检查"],
            ["建仓执行", "成交、风险暴露、资产偏离和交易成本复核"],
            ["日常管理", "净值与合成指数基准跟踪、策略单元归因、风格、信用、流动性和限额监控"],
            ["调仓处置", "调整理由、权限、成本及调整后风险复核"],
            ["产品复盘", "持有期结果、归因、客户体验和策略有效性评估"],
        ],
        [1800, 7560],
        font_size=9,
        first_col_bold=True,
    )
    add_callout(
        doc,
        "回撤管理原则",
        "不依赖单一机械止损。先识别回撤来自资产Beta、风格拥挤、信用、流动性还是管理人Alpha失效，再决定降风险顺序、流动性准备、渠道沟通和恢复条件。",
        fill=PALE_GOLD,
        accent=GOLD,
    )
    add_heading(doc, "把质量控制下沉到策略单元", 2)
    add_para(doc, "产品可以归因到每一个策略“积木块”，将组合收益与风险拆分为资产配置、策略选择、标的选择、交易与费用等来源；同时将实际净值与由SAA目标权重和代理指数构成的合成指数基准持续对比，确保产品在运作中没有悄然改变性格。")
    add_table(
        doc,
        ["监控对象", "核心指标", "出现偏离后的诊断"],
        [
            ["策略单元归因", "收益贡献、风险贡献、Alpha、下行贡献、费用贡献", "识别配置失效、策略失效或标的执行偏差"],
            ["实际净值 vs 合成基准", "累计差异、滚动超额、跟踪偏离、回撤差异", "区分资产偏离、风格漂移、现金拖累和交易成本"],
            ["产品特征", "高波动资产、久期、信用、风格因子、流动性和容量", "判断是否偏离产品矩阵单元的定义"],
            ["持有期体验", "正收益概率、尾部收益、最大回撤与修复时间", "评估是否需要调低风险、调整推荐持有期或加强陪伴"],
        ],
        [2050, 3350, 3960],
        font_size=8.3,
        first_col_bold=True,
    )
    add_figure(
        doc,
        examples["attribution"],
        "示例图：组合收益穿透至策略“积木块”的贡献归因",
        "模拟数据，仅用于展示归因层级和图表形式。",
        width=6.05,
    )
    add_para(doc, "示例阅读：若组合收益主要依赖单一策略，应进一步检查风险集中度；若某个“积木块”持续产生负Alpha，则需区分市场Beta、风格逆风、标的选择和费用拖累后再决定是否替换。", size=9.3, color=GRAY)
    add_asset_slot(
        doc, "A-03", "策略单元绩效与风险归因",
        "置于质量控制章节，展示组合结果如何穿透到每个“积木块”。",
        "采用日频费后净值；收益贡献与风险贡献分别计算；区分资产配置、策略选择、标的选择、交易与费用影响。",
        "收益贡献瀑布图＋风险贡献堆叠图＋策略单元预警清单。",
    )
    add_figure(
        doc,
        examples["tracking"],
        "示例图：实际净值与合成指数基准的跟踪及偏离诊断",
        "模拟数据；合成基准按SAA目标权重和代理指数构建。",
        width=6.05,
    )
    add_para(doc, "示例阅读：持续偏离不必然代表管理失效。偏离可能来自主动资产配置、标的Alpha、现金拖累、费用或执行时点；只有在归因后确认产品风格发生非预期漂移，才触发结构性调整。", size=9.3, color=GRAY)
    add_asset_slot(
        doc, "A-04", "实际净值与合成指数基准跟踪",
        "紧接归因图，展示产品特征管理的日常抓手。",
        "合成基准采用经审批的SAA目标权重与资产/策略代理指数，明确再平衡、费用、现金和基准切换规则。",
        "实际净值与合成基准累计曲线＋滚动超额/跟踪偏离子图＋偏离原因标注。",
    )
    add_placeholder(
        doc, "T-08", "历史压力测试与风险贡献拆解",
        "验证六个核心产品在典型市场冲击下的承受能力。",
        "选择具有代表性的股债、信用、流动性和商品冲击窗口，并对资产簇、策略簇和管理人贡献进行拆解。",
        "情景损益瀑布图、资产风险贡献图、流动性折损和恢复时间表。",
    )

    page_break(doc)
    add_heading(doc, "十三、渠道销售逻辑：先识别客户，再匹配矩阵单元", 1)
    add_table(
        doc,
        ["客户问题", "渠道判断", "匹配重点"],
        [
            ["资金多久可能使用？", "先确认最低可用期限和流动性需求", "选择最短持有期，不以收益倒推期限"],
            ["能承受多大净值波动？", "区分短期浮亏敏感度与长期风险承受能力", "选择风险风格和高波动资产预算"],
            ["更关心稳健还是弹性？", "识别客户对正收益概率、回撤和上涨弹性的排序", "选择防御、均衡或成长风格组合"],
            ["市场下跌会不会赎回？", "评估客户行为和陪伴需求", "使用推荐持有期和回撤情景说明"],
        ],
        [2500, 3550, 3310],
        font_size=8.8,
        first_col_bold=True,
    )
    add_heading(doc, "渠道标准表达", 2)
    add_bullets(doc, [
        "最短持有期：每笔份额在规定期限内不能赎回，满期后可在开放日按规则申请赎回。",
        "推荐持有期：根据资产配置、策略特征和历史测算建议的投资观察周期，不代表收益承诺。",
        "高波动资产：既包括股票、转债、商品和REITs等资产，也包括实际承担较高Beta、杠杆、波动率或流动性风险的策略。",
        "产品差异：不仅是权益比例不同，更是固收久期、资产簇、策略“积木块”、风格组合、风险预算和推荐持有期不同。",
    ])
    add_callout(
        doc,
        "渠道不应这样表达",
        "“持有满期一定盈利”“推荐持有期等于保本期限”“专户一定比公募收益高”“低波动等于无风险”“正收益概率代表未来保证”。",
        fill="FCEEEE",
        accent=RED,
    )

    page_break(doc)
    add_heading(doc, "十四、渠道合作价值：从单品销售走向产品矩阵经营", 1)
    add_table(
        doc,
        ["渠道价值", "具体体现"],
        [
            ["货架更清晰", "风险风格和期限两条轴线，减少同类产品堆叠"],
            ["客户匹配更准确", "先判断资金期限和风险承受能力，再进入具体产品"],
            ["销售语言更一致", "统一解释最短持有期、推荐持有期、高波动资产和产品差异"],
            ["客户陪伴更有依据", "用收益分布、回撤和修复周期替代单点收益解释"],
            ["产品复盘更可执行", "按矩阵单元评估规模、客户结构、策略容量和持有结果"],
            ["合作可持续", "建立产品、投研、风险和渠道反馈的闭环"],
        ],
        [2300, 7060],
        font_size=9,
        first_col_bold=True,
    )
    add_para(doc, "2025年末，31家理财公司的产品已打通母行之外的代销渠道，全市场跨行代销理财公司产品的机构达到593家。渠道合作正在从“有没有产品”转向“产品是否具有清晰定位、持续业绩和可解释的客户体验”。")
    add_source(doc, "《中国银行业理财市场年度报告（2025年）》。")
    add_callout(
        doc,
        "合作主张",
        "以产品矩阵为共同语言，以持有期数据为客户沟通依据，以全流程质量控制保障产品特征稳定，形成发行人与渠道共同经营客户长期体验的合作模式。",
    )

    page_break(doc)
    add_heading(doc, "十五、落地路径：先校准模型，再形成可发行、可销售、可维护的产品", 1)
    add_numbered(doc, [
        "确认首期核心产品单元、监管分类、正式命名和渠道优先级。",
        "建立资产与策略代理指数，完成SAA、费用和流动性假设。",
        "完成最短持有期、推荐持有期、正收益概率、回撤和压力测试。",
        "确定公募基金、专户及管理人的统一“积木块”分类、聚类与准入规则。",
        "形成产品参数卡、投资管理方案、风险评估和渠道路演材料。",
        "上线后持续跟踪客户结构、份额解锁、持有期结果和渠道反馈。",
    ])
    add_placeholder(
        doc, "T-09", "首期六类产品参数总表",
        "作为产品立项、渠道准入和后续销售培训的统一参数底稿。",
        "以最终审批口径汇总监管分类、SAA、TAA区间、费用、基准、风险指标、流动性和管理人限额。",
        "一页式产品参数矩阵＋六张单品卡。",
    )
    add_placeholder(
        doc, "T-10", "渠道客户适配与持有行为分析",
        "检验产品设计是否与真实渠道客户结构匹配。",
        "使用渠道客户风险等级、申购规模、持有时长、赎回时点和回撤期行为数据分组分析。",
        "客户分群、持有时长分布、回撤期赎回率及产品匹配建议。",
    )

    page_break(doc)
    add_heading(doc, "附录一：待测算图表清单", 1)
    add_table(
        doc,
        ["编号", "图表主题", "主要用途"],
        [
            ["T-01", "期限梯度与客户持有体验", "比较不同期限下正收益和尾部风险变化"],
            ["T-02", "完整产品矩阵风险收益地图", "展示矩阵单元的风险收益定位"],
            ["T-03", "六类产品持有期画像表", "形成渠道产品比较页"],
            ["T-04", "高波动资产风格指纹", "展示风格、行业和因子差异"],
            ["T-05", "公募基金与专户实施效率比较", "选择更合适的实施载体"],
            ["T-06", "市场状态与TAA决策面板", "解释动态调整规则"],
            ["T-07", "最短与推荐持有期双窗口测算", "支撑期限设计和客户陪伴"],
            ["T-08", "历史压力测试与风险贡献拆解", "验证极端环境承受能力"],
            ["T-09", "首期六类产品参数总表", "统一立项和渠道参数"],
            ["T-10", "渠道客户适配与持有行为分析", "检验产品与客户结构匹配"],
        ],
        [900, 3900, 4560],
        font_size=8.8,
        first_col_bold=True,
    )
    add_para(doc, "所有待测算图表应使用统一数据版本、费用假设、再平衡规则和样本区间；历史结果和模拟结果须明确区分，并在对外使用前完成投资、风险和合规复核。", size=9.5, color=GRAY)
    add_heading(doc, "已有图表素材的嵌入清单", 2)
    add_table(
        doc,
        ["编号", "素材主题", "嵌入前统一口径"],
        [
            ["A-01", "统一分类与策略聚类", "同观察期、同频率、费后净值、统一风险因子与人工复核"],
            ["A-02", "典型组合构成与回测", "净值归一、费用、再平衡、现金拖累、实盘/回测区分"],
            ["A-03", "策略单元归因", "收益贡献与风险贡献分列，归因层级与产品账簿一致"],
            ["A-04", "实际净值与合成指数基准", "SAA权重、代理指数、费用、现金及基准切换规则一致"],
        ],
        [900, 3600, 4860],
        font_size=8.6,
        first_col_bold=True,
    )

    page_break(doc)
    add_heading(doc, "附录二：主要资料来源", 1)
    sources = [
        "银行业理财登记托管中心：《中国银行业理财市场年度报告（2025年）》，2026年1月。",
        "中国证监会：《公开募集证券投资基金运作指引第2号——基金中基金指引》。",
        "国家金融监督管理总局及原中国银保监会：商业银行理财业务、理财公司销售及流动性风险管理相关制度。",
        "招商银行代销页面：信银理财安盈象固收稳利一个月持有期101号理财产品。",
        "中银理财：《中银理财－悦享（90天持有期）产品说明书》。",
        "工银理财：《核心优选最短持有180天固定收益类开放式理财产品（26GS5933）发行公告》。",
        "中邮理财：《灵活添利·鸿锦最短持有365天8号理财产品发行公告》。",
        "中国人民银行：《中国金融稳定报告（2019）》，关于2018年股票市场压力与主要指数表现。",
        "中国人民银行：《中国金融稳定报告（2021）》，关于2020年疫情冲击与金融稳定评估。",
        "中国人民银行：《中国金融稳定报告（2023）》相关章节，关于2022年11月债市调整与理财赎回反馈。",
        "华泰柏瑞中证500增强策略ETF 2024年第一季度报告、交银施罗德智选星光FOF 2024年中期报告，关于2024年初市场风格分化。",
        "上海证券交易所转载国新办2024年9月24日金融支持经济高质量发展新闻发布会材料。",
        "《泛FOF型多资产理财产品规划书》V0.2，2026年7月。",
        "已有内部策略分类、组合回测与固收专户比较材料，数据区间以各图表标注为准。",
    ]
    for idx, source in enumerate(sources, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.375)
        p.paragraph_format.first_line_indent = Inches(-0.194)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(f"{idx}.  {source}")
        set_run_font(r, size=9.5, color=INK)
    add_para(doc, "数据使用说明：公开市场数据截至各资料披露日；产品案例仅用于说明市场实践，不构成对相关产品的推荐或评价。最终产品要素以正式审批文件、产品合同和销售文件为准。", size=9, color=GRAY, italic=True)

    # Core properties and save
    doc.core_properties.title = "泛FOF型多资产理财产品方案——渠道路演材料"
    doc.core_properties.subject = "面向银行渠道产品经理与销售经理的产品方案"
    doc.core_properties.author = "产品规划项目组"
    doc.core_properties.keywords = "泛FOF, 银行理财, 产品矩阵, 持有期, 渠道路演"
    doc.core_properties.comments = "正式渠道交流稿；含待测算图表及已有素材嵌入标准占位。"
    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
