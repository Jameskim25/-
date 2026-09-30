# -*- coding: utf-8 -*-
"""SK trichem 양식(첨부 원본 PPT의 마스터·제목·로고·Confidential·페이지 번호)을 유지하고
본문을 편집 가능한 도형·표·차트로 다시 구성한다. 모든 수치는 재계산된 Excel 값을 읽어 사용한다."""
import copy
import json
import sys

import openpyxl
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SRC, XLSX, OUT = sys.argv[1:4]
MAP = json.load(open(XLSX + ".map.json"))
wb = openpyxl.load_workbook(XLSX, data_only=True)
S = wb["06_Report_Summary"]; M = wb["04_Monthly_2026_2027"]; B = wb["02_Batch_Raw"]; I = wb["01_Inputs"]


def K(name):
    return S.cell(row=MAP["KPI"][name], column=2).value


def f0(x):
    return f"{x:,.0f}"


def f1(x):
    return f"{x:,.1f}"


def f2(x):
    return f"{x:,.2f}"


def t2(kg):
    return f"{kg/1000:,.2f}"


FONT = "Noto Sans KR"
RED, ORG, DARK, GRAY, LG = "EB002C", "FF7900", "262626", "7F7F7F", "D9D9D9"
BLUE, GREEN, CX = "2E75B6", "548235", "2E75B6"
TINT = {"raw": ("F2F2F2", "7F7F7F"), "ref": ("FFE9D6", "FF7900"), "qc": ("DEEBF7", "2E75B6"),
        "tank": ("FFFFFF", "7F7F7F"), "fill": ("E2F0D9", "548235"), "out": ("FFFFFF", "404040")}
UNK = "FFF2CC"

prs = Presentation(SRC)
KEEP = {"TextBox 2", "직사각형 3", "TextBox 5", "타원 6", "그림 7", "직사각형 160", "직선 연결선[R] 162", "TextBox 10", "TextBox 163"}


# ------------------------------------------------------------------ helpers
def rgb(h):
    return RGBColor.from_string(h)


def set_run(run, size=8, bold=False, color=DARK, italic=False):
    f = run.font
    f.size = Pt(size); f.bold = bold; f.italic = italic; f.color.rgb = rgb(color); f.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", FONT)


def add_text(slide, x, y, w, h, parts, size=8, align="l", anchor="t", color=DARK, bold=False, margin=0.02,
             fill=None, line=None, shape=None, lw=0.75, dash=False, spacing=None):
    """parts: str | list of paragraphs; paragraph = str | list of (text, {opts})"""
    if shape is None:
        sh = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = 0.12
        sh.shadow.inherit = False
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = rgb(fill)
    elif shape is not None:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line); sh.line.width = Pt(lw)
        if dash:
            sh.line.dash_style = MSO_LINE.DASH
    elif shape is not None:
        sh.line.fill.background()
    tf = sh.text_frame
    tf.word_wrap = True
    m = Inches(margin)
    tf.margin_left = tf.margin_right = m; tf.margin_top = tf.margin_bottom = Inches(0.01)
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(parts, str):
        parts = [parts]
    for i, para in enumerate(parts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        if spacing:
            p.space_after = Pt(spacing)
        runs = [(para, {})] if isinstance(para, str) else para
        for txt, o in runs:
            r = p.add_run(); r.text = txt
            set_run(r, o.get("size", size), o.get("bold", bold), o.get("color", color), o.get("italic", False))
    return sh


def rect(slide, x, y, w, h, fill=None, line=None, lw=0.75, dash=False, shape=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line); sh.line.width = Pt(lw)
        if dash:
            sh.line.dash_style = MSO_LINE.DASH
    else:
        sh.line.fill.background()
    return sh


def line(slide, x1, y1, x2, y2, color=GRAY, w=1.0, arrow=True, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color); c.line.width = Pt(w)
    if dash:
        c.line.dash_style = MSO_LINE.DASH
    if arrow:
        ln = c.line._get_or_add_ln()
        te = etree.SubElement(ln, qn("a:tailEnd")); te.set("type", "triangle"); te.set("w", "sm"); te.set("len", "sm")
    return c


def cell_border(cell, color="BFBFBF", w=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w=str(w), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill")); c = etree.SubElement(sf, qn("a:srgbClr")); c.set("val", color)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
    # fill must precede borders? schema: lnL,lnR,lnT,lnB, ... then fill -> move fill after borders
    for tag in ("a:solidFill", "a:noFill"):
        f = tcPr.find(qn(tag))
        if f is not None:
            tcPr.remove(f); tcPr.append(f)


def table(slide, x, y, colw, rowh, data, size=7, header_rows=1, fills=None, colors=None, bolds=None, aligns=None, merges=()):
    nr, nc = len(data), len(colw)
    gs = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(colw)), Inches(rowh * nr))
    tbl = gs.table
    tblPr = tbl._tbl.tblPr
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"   # No Style, Table Grid
    tbl.first_row = False; tbl.horz_banding = False
    for j, w in enumerate(colw):
        tbl.columns[j].width = Inches(w)
    for i in range(nr):
        tbl.rows[i].height = Inches(rowh)
        for j in range(nc):
            cell = tbl.cell(i, j)
            v = data[i][j]
            cell.margin_left = cell.margin_right = Inches(0.03)
            cell.margin_top = cell.margin_bottom = Inches(0.0)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            al = (aligns[i][j] if aligns and aligns[i] else None) or ("l" if j == 0 else "c")
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[al]
            r = p.add_run(); r.text = "" if v is None else str(v)
            isH = i < header_rows
            col = (colors[i][j] if colors and colors[i] and colors[i][j] else ("FFFFFF" if isH else DARK))
            b = (bolds[i][j] if bolds and bolds[i] is not None else None)
            set_run(r, size, isH or bool(b) or j == 0, col)
            fc = fills[i][j] if fills and fills[i] and fills[i][j] else ("404040" if isH else "FFFFFF")
            cell.fill.solid(); cell.fill.fore_color.rgb = rgb(fc)
            cell_border(cell)
    for (r0, c0, r1, c1) in merges:
        tbl.cell(r0, c0).merge(tbl.cell(r1, c1))
    return gs


def clean_slide(slide, title, subtitle, num):
    for sh in list(slide.shapes):
        if sh.name not in KEEP:
            sh._element.getparent().remove(sh._element)
    for sh in slide.shapes:
        if sh.name == "TextBox 2":
            set_text_keep(sh, title)
        elif sh.name == "TextBox 5":
            set_text_keep(sh, subtitle)
        elif sh.name == "TextBox 163":
            set_text_keep(sh, "노란 공란 : 미확인 — 담당부서 확인 후 입력")
    # drop unused image relationships (e.g. orphan image138)
    used = {el.get(qn("r:embed")) for el in slide._element.iter() if el.get(qn("r:embed"))}
    for rId, rel in list(slide.part.rels.items()):
        if rel.reltype.endswith("/image") and rId not in used:
            slide.part.drop_rel(rId)


def set_text_keep(sh, text):
    p = sh.text_frame.paragraphs[0]
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    for extra in sh.text_frame.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)


def section(slide, x, y, text, w=5.0, sub=None):
    parts = [(text, {"bold": True, "size": 9, "color": DARK})]
    if sub:
        parts.append(("   " + sub, {"size": 7, "color": GRAY}))
    rect(slide, x, y + 0.035, 0.07, 0.13, fill=RED)
    add_text(slide, x + 0.12, y, w, 0.2, [parts], margin=0)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def pbox(slide, x, y, w, h, kind, head, sub=None, headsize=8, subsize=6.5, bold=True):
    f, l = TINT[kind]
    paras = [[(head, {"bold": bold, "size": headsize, "color": DARK})]]
    if sub:
        for s_ in (sub if isinstance(sub, list) else [sub]):
            paras.append([(s_, {"size": subsize, "color": "404040"})])
    return add_text(slide, x, y, w, h, paras, align="c", anchor="m", fill=f, line=l, shape=MSO_SHAPE.ROUNDED_RECTANGLE, lw=1.0)


# ------------------------------------------------------------------ process flow (slides 1·2)
def flow(slide, tobe=False):
    y1, h1 = 1.15, 0.62
    xs = [0.45 + k * 1.53 for k in range(6)]
    w = 1.23
    items = [("raw", "레이크 CpZr Crude", "입고"), ("qc", "IQC", "수입검사 · NMR·ICP·IC"), ("raw", "준비 · 투입", "Canister → 정제기"),
             ("ref", "공용 정제기 1대", "리플럭스 적용 ('27.7~)" if tobe else "103℃ · 57 h/Batch"),
             ("qc", "PQC", "공정검사 · NMR"), ("tank", "Product Tank", "본류 저장")]
    for k, (kind, a, b) in enumerate(items):
        if kind == "ref":
            pbox(slide, xs[k] - 0.05, y1 - 0.06, w + 0.1, h1 + 0.12, kind, a, [b, "개선 영향 구간 · 증설 아님" if tobe else "하이닉스·CXMT·이지켐 공동 사용"], headsize=8.5, subsize=6.5)
        else:
            pbox(slide, xs[k], y1, w, h1, kind, a, b)
        if k < 5:
            x0 = xs[k] + w + (0.05 if kind == "ref" else 0)
            line(slide, x0 + 0.02, y1 + h1 / 2, xs[k + 1] - (0.07 if items[k + 1][0] == "ref" else 0.02), y1 + h1 / 2)
    # Tank -> FQC elbow
    tx = xs[5] + w / 2
    yb = y1 + h1; ymid = 2.0
    line(slide, tx, yb, tx, ymid, arrow=False)
    line(slide, tx, ymid, 1.06, ymid, arrow=False)
    line(slide, 1.06, ymid, 1.06, 2.12)
    # row 2
    y2 = 2.12
    pbox(slide, 0.45, y2, 1.23, 0.84, "qc", "FQC", ["제품검사", "NMR·ICP-MS·IC·점도"])
    lanes = [(2.12, "5 Gal 글로브 박스 충진", "5 Gal(약 19 L) · 20 kg/병 · 필터 포함" + (" · 유지" if tobe else ""), "■ SK하이닉스  ■ CXMT", [RED, CX]),
             (2.56, "200 L ARS 자동 충진" if tobe else "200 L 수동 충진", ("140 kg/용기 · '26 설치 · '27 운영 · 필터 포함" if tobe else "140 kg/용기 · 글로브 박스 미사용 · 필터 포함"), "■ 이지켐", [ORG])]
    for (yy, a, b, cust, ccols) in lanes:
        line(slide, 1.68, y2 + 0.42, 1.98, yy + 0.2)
        pbox(slide, 2.0, yy, 2.55, 0.4, "fill", a, b, headsize=7.5, subsize=6)
        line(slide, 4.57, yy + 0.2, 4.83, yy + 0.2)
        pbox(slide, 4.85, yy, 0.95, 0.4, "qc", "OQC · 출하", None, headsize=7.5)
        line(slide, 5.82, yy + 0.2, 6.05, yy + 0.2)
        parts = []
        for tok in cust.split("  "):
            sq, name = tok[0], tok[1:]
            parts.append((sq, {"color": ccols[len(parts) // 2], "size": 8}))
            parts.append((name + "  ", {"bold": True, "size": 7.5}))
        add_text(slide, 6.07, yy + 0.08, 1.45, 0.25, [parts], margin=0)
    # parallel callout
    txt = [[("병렬 운영", {"bold": True, "color": RED, "size": 8})],
           [("앞 Batch 검사·충진 중 공용 정제기는 다음 Batch 정제", {"size": 6.8})],
           [("→ 충진시간을 정제시간에 더하지 않음", {"size": 6.8, "bold": True})]]
    if tobe:
        txt = [[("To-be 구분 원칙", {"bold": True, "color": RED, "size": 8})],
               [("정제기 1대 공정·운전 개선 (증설 아님)", {"size": 6.8})],
               [("ARS 충진 개선 효과는 Capa.에 미가산", {"size": 6.8, "bold": True})]]
    add_text(slide, 7.62, 2.12, 1.93, 0.84, txt, anchor="m", fill="FDE9E7", line=RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE, lw=0.75, margin=0.06)


def time_table(slide, y, tobe=False):
    hdr = ["공정 / 항목", "입고·IQC", "준비·투입", "정제" + (" (리플럭스)" if tobe else ""), "PQC", "Tank 이송", "FQC", "5 Gal 충진", "200 L 충진" + (" (ARS)" if tobe else " (수동)"), "OQC·출하"]
    work = ["작업시간 (h)", "2", "", "" if tobe else "57", "", "", "2 (+대기 미확인)", "2/병 · 약 8병 16", "" if tobe else "8", "2"]
    unit = ["적용 단위", "원료 Lot", "Batch", "Batch", "Batch", "Batch", "Batch", "병 / Batch", "용기 or Batch 확인", "출하 Lot"]
    res = ["점유 자원", "QC", "정제기", "공용 정제기 1대", "QC", "Tank", "QC", "글로브 박스", "ARS" if tobe else "200 L 수동", "QC·물류"]
    data = [hdr, work, unit, res]
    colw = [0.95, 0.8, 0.8, 1.0, 0.72, 0.8, 0.95, 1.1, 1.08, 0.9]
    hf = ["404040", "7F7F7F", "7F7F7F", ORG, BLUE, "7F7F7F", BLUE, GREEN, GREEN, BLUE]
    fills = [hf, ["F2F2F2"] + [None] * 9, ["F2F2F2"] + [None] * 9, ["F2F2F2"] + [None] * 9]
    for j in range(1, 10):
        if work[j] == "":
            fills[1][j] = UNK
    if tobe:
        for j in (1, 2, 4, 5, 6, 7, 9):
            pass
    colors = [None, [None] + [RED if (j == 3 and not tobe) else None for j in range(1, 10)], None, None]
    table(slide, 0.45, y, colw, 0.21, data, size=6.8, fills=fills, colors=colors)
    note = ("※ 정제 57 h에 포함된 준비·냉각·배출 시간은 재가산하지 않음 · IQC는 원료 Lot 단위(여러 Batch 공통 시 중복 배정 안 함) · 필터는 충진에 포함 — 별도 시간 가산 없음"
            if not tobe else "※ 리플럭스 후 정제시간·Batch량과 ARS 충진시간은 미확인 — 현재 57 h로 47.2톤을 계산하지 않음 · 5 Gal 충진·검사 조건은 현재와 동일")
    add_text(slide, 0.45, y + 0.93, 9.1, 0.17, [[(note, {"size": 6.3, "color": GRAY})]], margin=0)


# ------------------------------------------------------------------ slide 1
s1, s2, s3, s4, s5 = prs.slides
clean_slide(s1, "As-is | 공정시간과 현재 생산능력", "공용 정제기 1대 · 103℃ · 57 h/Batch · 190 kg/Batch · 26.2톤/년", 1)
flow(s1)
section(s1, 0.45, 3.1, "공정시간 (제공값)", sub="미확인 항목은 노란 공란 — Excel 01_Inputs에서 입력")
time_table(s1, 3.35)
# bottom-left : Lead time vs interval
section(s1, 0.45, 4.62, "한 Batch Lead Time ≠ 정제 생산 간격")
lt5, lt2 = K("5 Gal 경로 확인 Lead Time"), K("200 L 경로 확인 Lead Time")
tiles = [(0.45, "① 제조 Lead Time", [[("5 Gal ", {"size": 7.5}), (f"{lt5:.0f} h", {"size": 15, "bold": True, "color": DARK}), ("  +미확인", {"size": 6.5, "color": GRAY})],
                                             [("200 L ", {"size": 7.5}), (f"{lt2:.0f} h", {"size": 15, "bold": True, "color": DARK}), ("  +미확인", {"size": 6.5, "color": GRAY})],
                                             [(f"IQC 2+정제 57+FQC 2+충진+OQC 2 · 미확인 {K('Lead Time 미확인 항목 수'):.0f}항목 제외", {"size": 6, "color": GRAY})]]),
         (2.02, "② 다음 Batch 간격", [[("이론 ", {"size": 7.5}), ("57 h", {"size": 15, "bold": True, "color": ORG}), ("  정제만", {"size": 6.5, "color": GRAY})],
                                          [("추정 ", {"size": 7.5}), (f"{K('관측 추정 간격'):.1f} h", {"size": 15, "bold": True, "color": RED}), ("  과거 일정", {"size": 6.5, "color": GRAY})],
                                          [(f"범위 {K('관측 추정 하한'):.1f}~{K('관측 추정 상한'):.1f} h (3장)", {"size": 6, "color": GRAY})]]),
         (3.59, "③ 후공정 (별도)", [[("5 Gal ", {"size": 7.5}), ("16 h", {"size": 15, "bold": True, "color": GREEN}), ("/Batch", {"size": 6.5, "color": GRAY})],
                                         [("200 L ", {"size": 7.5}), ("8 h", {"size": 15, "bold": True, "color": GREEN}), ("/용기·Batch 확인", {"size": 6.5, "color": GRAY})],
                                         [("글로브 박스·인력 근무시간은 별도 입력", {"size": 6, "color": GRAY})]])]
for x, head, body in tiles:
    add_text(s1, x, 4.87, 1.5, 1.33, [[(head, {"bold": True, "size": 7.5})]] + body, fill="FFFFFF", line="BFBFBF",
             shape=MSO_SHAPE.RECTANGLE, margin=0.07, spacing=1)
add_text(s1, 0.45, 6.27, 4.64, 0.45, [[("수정: ", {"bold": True, "color": RED, "size": 7}),
                                       ("기존 '정제·충진 순차 → 월 9~10 Batch'는 Lead Time을 월 시간으로 나눈 방식 → 폐기. ", {"size": 7}),
                                       ("정제기 능력은 ②로, 충진 가능 여부는 ③으로 따로 확인", {"size": 7, "bold": True})]],
         fill="F2F2F2", shape=MSO_SHAPE.RECTANGLE, margin=0.07, anchor="m")
# bottom-right: capa
section(s1, 5.3, 4.62, "현재 Capa. 26.2톤/년 — 환산 관계")
add_text(s1, 5.3, 4.87, 1.55, 1.85, [[("현재 연간 Capa.", {"bold": True, "size": 8})],
                                     [("26.2", {"size": 30, "bold": True, "color": RED}), (" 톤/년", {"size": 9, "bold": True, "color": RED})],
                                     [("제공 기준값 · 103℃", {"size": 6.5, "color": GRAY})],
                                     [("시간·가동률을 조정해 맞추지 않음", {"size": 6.5, "color": GRAY})]],
         align="c", anchor="m", fill="FDE9E7", line=RED, shape=MSO_SHAPE.RECTANGLE, margin=0.05)
cap_rows = [["환산 항목", "값", "산식"],
            ["Batch 상당", f"{K('26.2톤 Batch 상당'):.1f} Batch/년", "26,200 ÷ 190"],
            ["정수 Batch (반올림)", f"138 × 190 = {f0(K('26.2톤 반올림 정수 Batch 생산량'))} kg", "26.22톤 → 표시 26.2톤"],
            ["월 평균", f"{K('월 평균 Batch 상당 (26.2톤)'):.2f} Batch 상당", "실제 월 계획 아님"],
            ["역산 평균 간격", f"{K('26.2톤 역산 간격'):.1f} h/Batch", "8,760 h 연속 가정"],
            ["시간 기반 추정 (3장)", f"이론 {t2(K('이론 연 생산량 (달력)'))} · 11 Batch {t2(K('11 Batch 달력 연속'))} · 관측 {t2(K('관측 연 생산량'))} 톤", "190 kg/Batch"]]
table(s1, 6.95, 4.87, [0.8, 1.15, 0.65], 0.305, cap_rows, size=6.3,
      aligns=[None] + [["l", "c", "l"]] * 5)
notes(s1, f"""[1장 As-is — 공정시간과 현재 생산능력]
■ 운영 전제: CpZr 정제기 1대를 하이닉스·CXMT·이지켐이 공동 사용. 앞 Batch를 검사·충진하는 동안 정제기는 다음 Batch를 정제(병렬). 초류·By-product·Residue·Waste 분기와 별도 Filtering 단계는 삭제(필터는 충진에 포함, 별도 시간 가산 없음).
■ 공정시간 (제공): IQC 2 h(원료 Lot당), 정제 57 h/Batch, FQC 2 h, 5 Gal 충진 2 h/병(약 8병 16 h), 200 L 충진 8 h(용기당/Batch당 확인), OQC·출하 2 h(운송 제외). 준비·투입, PQC, Tank 이송, FQC 대기, 세척·전환은 미확인(공란).
■ ① 제조 Lead Time(확인분): 5 Gal 경로 {lt5:.0f} h, 200 L 경로 {lt2:.0f} h — 미확인 {K('Lead Time 미확인 항목 수'):.0f}개 항목 제외 하한값. 이 값은 다음 Batch 정제 시작 간격이 아님.
■ ② 정제 생산 간격: 이론 57 h(정제만), 과거 일정 추정 {K('관측 추정 간격'):.2f} h (범위 {K('관측 추정 하한'):.2f}~{K('관측 추정 상한'):.2f} h, 3장).
■ ③ 후공정: 5 Gal 16 h/Batch(약 8병), 200 L 8 h — 글로브 박스·인력 근무시간은 정제기 720 h와 별도로 입력(Excel 01_Inputs F).
■ 기존 '정제·충진 순차 진행 → 월 9~10 Batch' 계산은 폐기 (Lead Time을 월 가용시간으로 나누는 방식).
■ 26.2톤 환산: 26,200 ÷ 190 = {K('26.2톤 Batch 상당'):.3f} Batch 상당/년 → 반올림 138 × 190 = 26,220 kg(26.22톤, 표시 26.2톤) / 내림 137 × 190 = 26,030 kg. 월 평균 {K('월 평균 Batch 상당 (26.2톤)'):.2f} Batch 상당. 8,760 h 연속 가정 역산 간격 {K('26.2톤 역산 간격'):.2f} h/Batch.
■ 시간 기반 추정과 비교(3장): 이론 {t2(K('이론 연 생산량 (달력)'))}톤, 과거 최대 11 Batch {t2(K('11 Batch 달력 연속'))}톤, 과거 일정 관측 {t2(K('관측 연 생산량'))}톤. 26.2톤은 기준값으로 유지하고 산출 조건(가동일·Batch량·집계 기준) 차이를 확인.
■ 상세: Excel 01_Inputs, 03_Capacity_Model 1·5절.""")

# ------------------------------------------------------------------ slide 2
clean_slide(s2, "To-be | 리플럭스 적용 공정과 계획", "동일 정제기 1대 개선 · '27.7 적용 생산 · 47.2톤/년 (연간 환산)", 2)
flow(s2, tobe=True)
section(s2, 0.45, 3.1, "공정시간 (To-be)", sub="리플럭스 정제시간·ARS 충진시간 미확인")
time_table(s2, 3.35, tobe=True)
section(s2, 0.45, 4.62, "리플럭스 · ARS 일정")
# gantt grid (editable shapes)
gx, gy, lw_, cw = 0.45, 4.87, 1.42, 0.19
months = [f"{m}" for m in range(1, 13)] + [f"{m}" for m in range(1, 7)]
add_text(s2, gx + lw_, gy, cw * 12, 0.16, [[("2027년", {"bold": True, "size": 6.5, "color": "FFFFFF"})]], align="c", anchor="m", fill="404040", shape=MSO_SHAPE.RECTANGLE, margin=0)
add_text(s2, gx + lw_ + cw * 12, gy, cw * 6, 0.16, [[("2028년", {"bold": True, "size": 6.5, "color": "FFFFFF"})]], align="c", anchor="m", fill=GRAY, shape=MSO_SHAPE.RECTANGLE, margin=0)
for k, mm in enumerate(months):
    add_text(s2, gx + lw_ + cw * k, gy + 0.16, cw, 0.15, [[(mm, {"size": 6})]], align="c", anchor="m", fill="F2F2F2", line="D9D9D9", shape=MSO_SHAPE.RECTANGLE, lw=0.25, margin=0)
grows = [("리플럭스 공사", 1, 5, ORG, "공사 2~5월", False),
         ("시운전", 5, 6, "FFC000", "6월", False),
         ("적용 생산", 6, 18, RED, "생산 시작 '27.7", False),
         ("최초 12개월", 6, 18, "F4B6C2", "'27.7~'28.6 · 47.2톤 환산 기준", False),
         ("ARS 200 L 운영", 0, 18, GREEN, "'26 설치 → '27 운영 (Capa. 미가산)", False),
         ("105℃ 검토안", 12, 18, None, "2028 · 적용 시점 미정", True)]
for i, (lab, a, b, col, txt, dash) in enumerate(grows):
    yy = gy + 0.34 + i * 0.2
    add_text(s2, gx, yy, lw_ - 0.04, 0.2, [[(lab, {"bold": True, "size": 6.8})]], anchor="m", margin=0)
    rect(s2, gx + lw_, yy + 0.2, cw * 18, 0.005, fill="E7E6E6")
    if dash:
        add_text(s2, gx + lw_ + cw * a, yy + 0.02, cw * (b - a), 0.17, [[(txt, {"size": 6, "color": "7030A0", "bold": True})]], align="c", anchor="m",
                 line="7030A0", dash=True, shape=MSO_SHAPE.RECTANGLE, margin=0)
    else:
        tc = "FFFFFF" if col in (RED, ORG, GREEN) else DARK
        add_text(s2, gx + lw_ + cw * a, yy + 0.02, cw * (b - a), 0.17, [[(txt, {"size": 6, "color": tc, "bold": True})]], align="c", anchor="m",
                 fill=col, shape=MSO_SHAPE.RECTANGLE, margin=0)
add_text(s2, 0.45, 6.45, 4.8, 0.3, [[("※ 2027년 달력연도 생산량은 공사·시운전·안정화·실제 가동조건으로 별도 산정 (5장) — 공사 중 생산량은 0 또는 정상값으로 임의 입력하지 않음", {"size": 6.3, "color": GRAY})]], margin=0)
# roadmap
section(s2, 5.45, 4.62, "Capa. 로드맵 (연간 환산 기준)")
bars = [(26.2, "현재 (103℃)", "제공 기준값", RED, False), (47.2, "리플럭스 후", "'27.7~'28.6 12개월 환산", RED, False), (50.2, "105℃ 적용 시", "검토안 · 검증·승인 필요", None, True)]
bx0, by_base, bw = 5.55, 6.3, 0.9
for k, (val, lab, sub, col, dash) in enumerate(bars):
    hh = val / 50.2 * 1.0
    x = bx0 + k * 1.3
    if dash:
        rect(s2, x, by_base - hh, bw, hh, fill="F2E6F7", line="7030A0", dash=True)
    else:
        rect(s2, x, by_base - hh, bw, hh, fill=col if k == 1 else "F4B6C2")
    add_text(s2, x - 0.1, by_base - hh - 0.26, bw + 0.2, 0.24, [[(f"{val:.1f}", {"size": 12, "bold": True, "color": "7030A0" if dash else RED}), (" 톤/년", {"size": 7, "color": GRAY})]], align="c", margin=0)
    add_text(s2, x - 0.15, by_base + 0.02, bw + 0.3, 0.34, [[(lab, {"bold": True, "size": 7})], [(sub, {"size": 6, "color": GRAY})]], align="c", margin=0)
    if k < 2:
        add_text(s2, x + bw + 0.02, by_base - 0.55 - k * 0.1, 0.36, 0.2, [[("+21" if k == 0 else "+3", {"bold": True, "size": 8, "color": RED if k == 0 else "7030A0"})]], align="c", margin=0)
add_text(s2, 5.45, 6.66 - 0.0, 4.1, 0.2, [[("ARS 효과 미가산 · 47.2톤 ≠ 2027 실제 생산량 · 등가 간격 ", {"size": 6.3, "color": GRAY}),
                                           (f"{K('47.2톤 등가 간격'):.1f} h/Batch", {"size": 6.3, "bold": True, "color": DARK}),
                                           (" (190 kg 유지 시 역산)", {"size": 6.3, "color": GRAY})]], margin=0)
add_text(s2, 5.5, 4.86, 1.2, 0.62, [[("105℃ 검증 필요", {"bold": True, "size": 7, "color": "7030A0"})],
                                    [("열 안정성 · Dimer", {"size": 6.3})], [("Unknown impurity · 수율", {"size": 6.3})],
                                    [("미확정·승인·시점 미정", {"size": 6.3, "bold": True})]],
         fill="F2E6F7", shape=MSO_SHAPE.RECTANGLE, margin=0.05)
notes(s2, f"""[2장 To-be — 리플럭스 공정과 계획]
■ 공정: 동일 공용 정제기 1대에 리플럭스 적용(증설 아님). 5 Gal(하이닉스·CXMT)은 글로브 박스 충진 유지, 200 L(이지켐)은 ARS 자동 충진(2026년 내 설치, 2027년 운영 계획). 초류·Waste 분기와 별도 Filtering 삭제 — 필터는 충진에 포함.
■ 일정: 2027년 2~5월 공사, 6월 시운전, 7월 리플럭스 적용 생산 시작, 최초 12개월 2027.7~2028.6.
■ Capa.: 현재 26.2 + 리플럭스 21 = 47.2톤/년 (12개월 동일 조건 연간 환산). 평균 월 {K('47.2 월 환산'):.2f}톤 = 190 kg 유지 시 {K('47.2 Batch 상당/월'):.2f} Batch 상당/월. 190 kg·8,760 h 가정 역산 등가 간격 {K('47.2톤 등가 간격'):.2f} h/Batch (리플럭스 후 실제 정제시간 미확인 — 현재 57 h로 47.2톤을 계산하지 않음).
■ 2028 검토안: 103→105℃, 예상 +3톤 → 50.2톤/년(산술값), 목표 2028년 이후 50톤 수준. 미확정·검증 및 승인 필요·적용 시점 미정. 검증: 열 안정성, Dimer 형성, Unknown impurity 증가, 수율.
■ 구분 원칙: ARS 충진 개선 효과는 47.2톤·50.2톤에 가산하지 않음. 47.2톤은 연간 환산값으로 2027년 실제 생산량이 아님(5장).
■ 미확인: 리플럭스 후 정제시간·Batch량·안정화 기간, 공사 중 정제기 가동 여부, 시운전 양품, ARS 운영 개시일·충진시간 (Excel 01_Inputs D·E).""")

# ------------------------------------------------------------------ slide 3
clean_slide(s3, "과거 Batch 일정과 병렬 운영 생산능력 추정",
            f"2026.6~10 생산계획 #39~#92 · 추정 실효 간격 약 {K('관측 추정 간격'):.1f} h/Batch · 월 {K('관측 월 Batch 30일'):.1f}~{K('관측 월 Batch 31일'):.1f} Batch", 3)
section(s3, 0.45, 1.08, "과거 Batch 일정 분포", w=5.3, sub="표시일자 = 원자료 기재일 (투입·완료 미확정)")
# read raw
raw = []
r = 6
while B.cell(row=r, column=4).value is not None and isinstance(B.cell(row=r, column=4).value, (int, float)):
    raw.append((int(B.cell(row=r, column=4).value), B.cell(row=r, column=5).value))
    r += 1
gx0, gx1 = 1.08, 4.9
dayw = (gx1 - gx0) / 30
gy0 = 1.5
for d in (1, 5, 10, 15, 20, 25, 31):
    add_text(s3, gx0 + (d - 1) * dayw - 0.1, gy0 - 0.17, 0.2, 0.14, [[(str(d), {"size": 5.5, "color": GRAY})]], align="c", margin=0)
add_text(s3, 4.98, gy0 - 0.2, 0.42, 0.18, [[("번호", {"size": 6, "bold": True})]], align="c", margin=0)
add_text(s3, 5.38, gy0 - 0.2, 0.42, 0.18, [[("Total", {"size": 6, "bold": True})]], align="c", margin=0)
mrows = [(5, "5월 (일부)"), (6, "6월"), (7, "7월"), (8, "8월"), (9, "9월"), (10, "10월")]
for k, (mm, lab) in enumerate(mrows):
    yy = gy0 + k * 0.2
    add_text(s3, 0.45, yy, 0.6, 0.16, [[(lab, {"bold": True, "size": 6.5})]], anchor="m", margin=0)
    rect(s3, gx0 - 0.05, yy + 0.08, gx1 - gx0 + 0.1, 0.005, fill="D9D9D9")
    for n, d in raw:
        if d.month == mm:
            col = GRAY if mm == 5 else ORG
            add_text(s3, gx0 + (d.day - 1) * dayw - 0.066, yy + 0.015, 0.132, 0.132, [[(str(n), {"size": 4.5, "color": "FFFFFF", "bold": True})]],
                     align="c", anchor="m", fill=col, shape=MSO_SHAPE.OVAL, margin=0)
    sr = MAP["SUMROW"][str(mm)]
    cnt = B.cell(row=sr, column=3).value; tot = B.cell(row=sr, column=4).value
    diff = tot is not None and cnt != tot
    add_text(s3, 4.98, yy, 0.42, 0.16, [[(f"{cnt}" + ("*" if mm == 5 else ""), {"size": 7, "bold": True, "color": RED if diff else DARK})]], align="c", anchor="m", margin=0)
    add_text(s3, 5.38, yy, 0.42, 0.16, [[("–" if tot is None else f"{tot}", {"size": 7, "bold": True, "color": RED if diff else DARK})]], align="c", anchor="m", margin=0)
st = MAP["SUMROW"]["tot"]
add_text(s3, 0.45, gy0 + 1.22, 5.4, 0.28,
         [[(f"6~10월 번호 {B.cell(row=st, column=3).value}개 vs Total {B.cell(row=st, column=4).value} (빨강 = 불일치, 번호≠양품 완료) · * 5월 일부 · 월 경계 간격 모두 3일 → 월말 진행 Batch 다음 월 이월", {"size": 5.8, "color": GRAY})]], margin=0)
# chart: gap distribution
section(s3, 6.0, 1.08, "표시일자 간격 분포", w=3.5, sub=f"#39~#92 · {K('관측 간격 수'):.0f}개")
cd = CategoryChartData()
g0 = MAP["GD0"]
cats = [S.cell(row=g0 + j, column=1).value for j in range(5)]
vals = [S.cell(row=g0 + j, column=2).value for j in range(5)]
cd.categories = cats
cd.add_series("건수", vals)
gf = s3.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.0), Inches(1.3), Inches(3.55), Inches(1.28), cd)
ch = gf.chart
ch.has_legend = False; ch.has_title = False
pl = ch.plots[0]; pl.gap_width = 60; pl.has_data_labels = True
pl.data_labels.font.size = Pt(7); pl.data_labels.font.bold = True; pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
pl.data_labels.number_format = "0"; pl.data_labels.number_format_is_linked = False
ser = pl.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = rgb(ORG)
ch.value_axis.visible = False; ch.value_axis.has_major_gridlines = False
ch.value_axis.maximum_scale = 46
ch.category_axis.tick_labels.font.size = Pt(7); ch.category_axis.format.line.color.rgb = rgb("BFBFBF")
ch.font.name = FONT
add_text(s3, 6.0, 2.58, 3.55, 0.3, [[("D일 간격의 실제 경과 = (D−1)×24 ~ (D+1)×24 h → ", {"size": 6, "color": GRAY}), ("2일 = 24~72 h", {"size": 6, "bold": True}),
                                     (" · 2일을 48 h로 단정해 57 h 불가로 결론 내리지 않음", {"size": 6, "color": GRAY})]], margin=0)
# gantt
section(s3, 0.45, 2.98, "병렬 운영 Gantt", w=6, sub=f"실효 간격 {K('관측 추정 간격'):.1f} h 기준 · 검사·충진이 다음 Batch 정제와 겹침")
tx0, tx1, hmax = 1.62, 9.5, 216
sc = (tx1 - tx0) / hmax
ty = 3.2
for hh in range(0, 217, 24):
    add_text(s3, tx0 + hh * sc - 0.18, ty, 0.36, 0.12, [[(f"{hh}" + (" h" if hh == 216 else ""), {"size": 5.5, "color": GRAY})]], align="c", margin=0)
    rect(s3, tx0 + hh * sc, ty + 0.13, 0.004, 0.83, fill="E7E6E6")
gint = round(K("관측 추정 간격"), 1)
lanes = ["공용 정제기 1대", "검사 (FQC·OQC)", "5 Gal 충진 (글로브 박스)"]
for i, lab in enumerate(lanes):
    add_text(s3, 0.45, ty + 0.15 + i * 0.27, 1.15, 0.22, [[(lab, {"bold": True, "size": 6.5})]], anchor="m", margin=0)
for b in range(3):
    st_ = b * gint
    def bar(lane, a, bb, col, txt, tc="FFFFFF"):
        a = min(a, hmax); bb = min(bb, hmax)
        if bb <= a:
            return
        add_text(s3, tx0 + a * sc, ty + 0.17 + lane * 0.27, (bb - a) * sc, 0.19, [[(txt, {"size": 5.5, "color": tc, "bold": True})]],
                 align="c", anchor="m", fill=col, shape=MSO_SHAPE.RECTANGLE, margin=0)
    bar(0, st_, st_ + 57, RED, f"Batch {b+1} 정제 57 h")
    bar(0, st_ + 57, st_ + gint, "BFBFBF", "", DARK)
    bar(1, st_ + 57, st_ + 59, BLUE, "")
    bar(2, st_ + 59, st_ + 75, GREEN, f"B{b+1} 충진 16 h")
    bar(1, st_ + 75, st_ + 77, BLUE, "")
add_text(s3, 0.45, ty + 0.95, 9.1, 0.14, [[("회색 = 비정제 추정 약 ", {"size": 5.8, "color": GRAY}), (f"{K('관측 비정제 시간'):.1f} h", {"size": 5.8, "bold": True}),
                                                          (" (준비·세척·대기 등, 구성 미확인) · 파랑 = FQC 2 h / OQC 2 h", {"size": 5.8, "color": GRAY})]], margin=0)
# comparison table
section(s3, 0.45, 4.3, "생산능력 비교 (정제기 기준 · 190 kg/Batch)", w=5.5)
c0 = MAP["CMP0"]; Cm = wb["03_Capacity_Model"]
def cm(k, col):
    return Cm.cell(row=c0 + k, column=col).value
rows = [["구분", "Batch 간격", "월 Batch (30일)", "연 Batch", "연 생산량", "26.2 대비"]]
labels = ["① 이론 (정제 57 h만)", "현재 Capa. 26.2톤 역산", "② 과거 월 최대 11 Batch", "③ 과거 일정 관측 추정", "참고: 원자료 Total 기준"]
extra = [f"{cm(0,3):.1f} (정수 12·{K('이론 월말 이월 h'):.0f} h 이월)", f"{cm(1,3):.1f}", f"11 (정제 {K('11 Batch 정제시간'):.0f} h·차이 {K('11 Batch 720 h 차이'):.0f} h)",
         f"{cm(3,3):.1f} (31일 {cm(3,4):.1f})", f"{cm(4,3):.1f}"]
for k in range(5):
    intv = f"{cm(k,2):.1f} h" + (f" ({K('관측 추정 하한'):.1f}~{K('관측 추정 상한'):.1f})" if k == 3 else "")
    by = f"{cm(1,5):.1f}" if k == 1 else f"{cm(k,5):.0f}"
    by = f"{K('26.2톤 Batch 상당'):.1f}" if k == 1 else by
    ty_ = "26.20 톤" if k == 1 else f"{cm(k,6):.2f} 톤"
    dd = "기준" if k == 1 else f"{cm(k,7):+.2f}"
    rows.append([labels[k], intv, extra[k], by, ty_, dd])
fills = [None] + [None, None, None, [ "FDE9E7"] * 6, ["F2F2F2"] * 6]
fills = [None, None, ["F2F2F2"] * 6, None, ["FDE9E7"] * 6, ["F2F2F2"] * 6]
colors = [None, None, None, None, [RED] * 6, [GRAY] * 6]
bolds = [None, None, None, None, [True] * 6, None]
table(s3, 0.45, 4.55, [1.45, 1.08, 1.47, 0.5, 0.66, 0.5], 0.235, rows, size=6.3, fills=fills, colors=colors, bolds=bolds)
add_text(s3, 0.45, 5.99, 5.66, 0.72, [[("판단: ", {"bold": True, "size": 6.8, "color": RED}),
                                       (f"현재 운영에서 정제기 실효 생산 간격은 약 {K('관측 추정 간격'):.1f} h/Batch(정제 57 h + 비정제 약 {K('관측 비정제 시간'):.1f} h)로 추정. "
                                        f"31일 월 약 11 Batch = 과거 최대 11 Batch와 정합. 연 {K('관측 연 Batch'):.0f} Batch·{t2(K('관측 연 생산량'))}톤은 26.2톤보다 {abs(K('관측 연 생산량')/1000-26.2):.2f}톤 낮음 → 26.2톤 산출 조건 확인 필요.", {"size": 6.8})],
                                      [("※ 연 Batch는 8,760 h 달력 연속(월간 이월 반영) 정수값, 정지·보수 미반영 · 단순 12배 참고: 이론 144, 11 Batch 132 Batch", {"size": 6, "color": GRAY})]],
         fill="F2F2F2", shape=MSO_SHAPE.RECTANGLE, margin=0.06, anchor="m")
# assumptions box
section(s3, 6.3, 4.3, "추정 근거 · 가정 · 확인사항", w=3.3)
ab = [("데이터", f"2026.6.2~10.29 표시일자, #39~#92 (간격 {K('관측 간격 수'):.0f}개, {K('관측 경과일수'):.0f}일)"),
      ("방법", f"경과일×24 ÷ 간격 수 → 장기 누적으로 날짜 오차 ±{(K('관측 추정 상한')-K('관측 추정 간격')):.2f} h"),
      ("분포", f"3일 간격 {K('관측 3일 간격 비중')*100:.0f}% · 중앙값 {K('관측 중앙값 h'):.0f} h · #37 포함 시 {K('#37 포함 평균'):.1f} h"),
      ("가정", "표시일자 = 매 Batch 동일 이벤트 · 번호 Batch 모두 수행"),
      ("적용 범위", "현재 103℃ 조건 한정 — 리플럭스 후 적용 불가"),
      ("확인", "표시일자 의미(투입/완료) · 실제 시각 · 계획 대비 실적 · Total 46 vs 번호 54 · 색상 의미"),
      ("정합 확인", "6/23~7/3 · 9/3~9/11: 4간격 평균 최대 54 h < 57 h → 계획 변경·표시 기준 확인")]
paras = [[("추정 실효 간격  ", {"bold": True, "size": 8}), (f"{K('관측 추정 간격'):.1f} h/Batch", {"bold": True, "size": 14, "color": RED}),
          (f"  ({K('관측 추정 하한'):.1f}~{K('관측 추정 상한'):.1f} h)", {"size": 7, "color": GRAY})]]
for a, b_ in ab:
    paras.append([(a + "  ", {"bold": True, "size": 6.8, "color": RED if a == "정합 확인" else DARK}), (b_, {"size": 6.8})])
add_text(s3, 6.3, 4.55, 3.25, 2.16, paras, fill="FFFFFF", line="BFBFBF", shape=MSO_SHAPE.RECTANGLE, margin=0.07, spacing=3.5, anchor="m")
notes(s3, f"""[3장 과거 Batch 일정과 병렬 운영 생산능력 추정]
■ 원자료(첨부 PPT 3장 발표자 노트 '원자료 5개 이미지 연결, 중복 제거' 전사 + 3장 분포도 좌표 판독 + 사용자 지시 목록 3중 대조 → 차이 {K('대조 차이 건수'):.0f}건). 원본 이미지는 첨부되지 않아 이미지 대조는 추가 확인사항. 연도 미표기 → 2026년 가정.
- 5월(일부): 5/23 #37, 5/26 #38
- 6월: 6/2 #39, 6/5 #40, 6/8 #41, 6/11 #42, 6/14 #43, 6/17 #44, 6/20 #45, 6/23 #46, 6/25 #47, 6/27 #48, 6/28 #49
- 7월: 7/1 #50, 7/3 #51, 7/6 #52, 7/8 #53, 7/11 #54, 7/14 #55, 7/17 #56, 7/20 #57, 7/23 #58, 7/26 #59, 7/29 #60
- 8월: 8/1 #61 ~ 8/31 #71 (3일 간격 11개)
- 9월: 9/3 #72, 9/5 #73, 9/7 #74, 9/9 #75, 9/11 #76, 9/14 #77, 9/19 #78, 9/23 #79, 9/26 #80, 9/28 #81
- 10월: 10/1 #82, 10/4 #83, 10/7 #84, 10/10 #85, 10/12 #86, 10/15 #87, 10/18 #88, 10/20 #89, 10/23 #90, 10/26 #91, 10/29 #92
■ 집계: 번호 개수 6월 11 / 7월 11 / 8월 11 / 9월 10 / 10월 11 = 54, 원자료 Total 8 / 11 / 8 / 8 / 11 = 46 (6·8·9월 불일치). 번호 개수를 양품 완료 Batch 수로 확정하지 않음.
■ 간격(#39~#92, 53개): 1일 1, 2일 11, 3일 39, 4일 1, 5일 1 (#37 포함 전체 55개: 3일 40, 2일 11, 1·4·5·7일 각 1). 월별 평균 6월 2.60 / 7월 2.82 / 8월 3.00 / 9월 2.80 / 10월 2.82일, 중앙값 3일(9월 2.5일).
■ 추정 방법: 6/2(#39)→10/29(#92) 149일, 53간격 → 149×24÷53 = {K('관측 추정 간격'):.2f} h/Batch. 시각 미상이므로 (149±1)×24÷53 = {K('관측 추정 하한'):.2f}~{K('관측 추정 상한'):.2f} h. 정제 57 h와의 차이 {K('관측 비정제 시간'):.2f} h/Batch는 준비·투입, 세척·전환, 검사·Tank 대기, 계획 여유 등의 가능성 — 구성 미확인, 전부 손실·충진시간으로 단정하지 않음.
■ 57 h 정합 확인: 개별 2일 간격은 24~72 h로 57 h와 양립 가능. 다만 연속 4간격(#46~#50/#51, #72~#76)은 8일(최대 216 h ÷ 4 = 54 h < 57 h) → 표시일자가 동일 이벤트라면 57 h 연속 정제와 맞지 않음 → 표시일자 의미·계획 변경·실적 여부 확인 (결론 보류).
■ 이론: 720 ÷ 57 = {K('이론 Batch/월 (소수)'):.2f} → 월내 완료 12 Batch(2,280 kg), 월말 {K('이론 월말 이월 h'):.0f} h 진행분 이월. 연간 달력 연속 INT(8,760÷57) = 153 Batch = 29,070 kg (단순 12×12 = 144 Batch = 27,360 kg 참고).
■ 과거 최대 11 Batch: 정제 627 h, 720 h와 차이 93 h, 배분 실효 65.45 h (차이 8.45 h/Batch; 31일 월이면 67.64 h). 월 2,090 kg, 단순 12배 25,080 kg, 달력 연속 133 Batch = 25,270 kg. 차이는 여유·전환·대기·보수·월 경계 가능성.
■ 관측 추정: 30일 월 {K('관측 월 Batch 30일'):.2f}, 31일 월 {K('관측 월 Batch 31일'):.2f} Batch → 11 Batch 월(7·8·10월)이 모두 31일 월인 점과 정합. 연 {K('관측 연 Batch'):.0f} Batch = {f0(K('관측 연 생산량'))} kg. 원자료 Total을 완료 Batch로 해석하면 {K('Total 기준 간격'):.1f} h/Batch(참고).
■ Gantt: 앞 Batch의 FQC 2 h → 5 Gal 충진 16 h → OQC 2 h가 다음 Batch 정제 57 h와 겹침 → 충진시간을 정제시간에 더해 능력을 줄이지 않음. Product Tank 1기이면 충진이 다음 Batch 정제 종료 전 끝나야 함(1교대 충진 경과 48 h → 57 h 대비 여유 {K('Tank 1교대 여유 vs 57'):.0f} h).
■ 상세: Excel 02_Batch_Raw, 03_Capacity_Model 2~9절.""")

# ------------------------------------------------------------------ slide 4
clean_slide(s4, "2026년 월별 생산계획 · 출하 비교",
            f"제시 출하 {t2(K('2026 제시 출하'))}톤 (CXMT 미제시) · 현재 Capa. 26.2톤 대비 {K('2026 제시 출하')/26200*100:.1f}%", 4)
section(s4, 0.45, 1.08, "2026년 월별 출하 · 과거 생산계획 · 생산 가능량", w=8, sub=f"단위 kg · 생산 가능량 = 실효 간격 {K('현재 조건 적용 간격'):.1f} h 기준 추정")
m0 = MAP["M0"]
hdr = ["구분"] + [f"{m}월" for m in range(1, 13)] + ["합계"]
def mrow(col, fmt=f0, rng=range(12)):
    out = []
    for k in rng:
        v = M[f"{col}{m0+k}"].value
        out.append(fmt(v) if isinstance(v, (int, float)) else ("" if v in (None, "") else str(v)))
    return out
hx = mrow("T"); ez = mrow("V"); tot = mrow("W")
cnt = mrow("L"); tt = mrow("M"); cap = mrow("S"); diff = mrow("AH", fmt=lambda v: f"{v:+,.0f}")
plan_num = [(f"{int(v)*190:,}" if v not in ("", None) else "") for v in cnt]
plan_tot = [(f"{int(v)*190:,}" if v not in ("", None) else "") for v in tt]
sr26 = MAP["SR"]["2026"]
rows4 = [hdr,
         ["■ SK하이닉스 출하"] + hx + [f0(M[f"T{sr26}"].value)],
         ["■ 이지켐 출하"] + ez + [f0(M[f"V{sr26}"].value)],
         ["■ CXMT 출하"] + ["미제시"] * 12 + ["미제시"],
         ["출하 합계 (제시분)"] + tot + [f0(M[f"W{sr26}"].value)],
         ["과거계획 번호×190"] + [p + ("*" if k == 4 and p else "") for k, p in enumerate(plan_num)] + ["10,260*"],
         ["과거계획 Total×190"] + plan_tot + ["8,740*"],
         ["생산 가능량 (추정)"] + cap + [f0(M[f"S{sr26}"].value)],
         ["생산 가능 − 출하"] + diff + [f"{M[f'AI{m0+11}'].value:+,.0f}"]]
fills = [None, None, None, [None] + ["F2F2F2"] * 13, ["F2F2F2"] * 14, None, None, ["FDE9E7"] * 14, None]
colors = [None, [RED] + [None] * 13, [ORG] + [None] * 13, [CX] + [GRAY] * 13, None, [GRAY] + [None] * 13, [GRAY] + [None] * 13, [RED] + [None] * 13,
          [None] + [RED if d.startswith("-") else GREEN for d in diff] + [None]]
bolds = [None, None, None, None, [True] * 14, None, None, [True] * 14, [True] * 14]
# highlight plan-number vs total mismatch months
for k in range(12):
    if plan_num[k] and plan_tot[k] and plan_num[k] != plan_tot[k]:
        colors[6] = colors[6] or [None] * 14
        colors[6][k + 1] = RED
table(s4, 0.45, 1.32, [1.2] + [0.555] * 12 + [0.8], 0.215, rows4, size=6.3, fills=fills, colors=colors, bolds=bolds)
add_text(s4, 0.45, 3.46, 9.1, 0.16, [[("* 합계는 6~10월 · 5월은 #37·#38 일부만 제공 · 과거계획은 생산계획(검증 실적 아님) · 빨간 Total = 번호 개수와 불일치 · 생산 가능량은 정제기 능력 추정(실제 계획·실적 아님) · 하이닉스 실적/계획 구분 미확인", {"size": 6, "color": GRAY})]], margin=0)
# chart
section(s4, 0.45, 3.66, "월별 출하 vs 생산 가능량 · 과거계획", w=5.4, sub="kg")
cd = CategoryChartData()
cd.categories = [f"{m}월" for m in range(1, 13)]
def num(col):
    return [M[f"{col}{m0+k}"].value if isinstance(M[f"{col}{m0+k}"].value, (int, float)) else None for k in range(12)]
cd.add_series("출하 합계 (제시분)", num("W"))
cd.add_series("생산 가능량 (추정)", num("S"))
cd.add_series("과거계획 Total×190", [(v * 190 if isinstance(v, (int, float)) else None) for v in num("M")])
gf = s4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.45), Inches(3.86), Inches(5.45), Inches(2.17), cd)
ch = gf.chart; ch.font.name = FONT; ch.font.size = Pt(7)
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout = False; ch.legend.font.size = Pt(6.5)
pl = ch.plots[0]; pl.gap_width = 50; pl.overlap = -10
for ser, col in zip(pl.series, [RED, "BFBFBF", ORG]):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = rgb(col)
ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = rgb("E7E6E6")
ch.value_axis.tick_labels.font.size = Pt(6.5); ch.value_axis.tick_labels.number_format = "#,##0"; ch.value_axis.tick_labels.number_format_is_linked = False
ch.value_axis.format.line.fill.background(); ch.value_axis.maximum_scale = 2500
ch.category_axis.tick_labels.font.size = Pt(6.5); ch.category_axis.format.line.color.rgb = rgb("BFBFBF")
# KPI tiles
section(s4, 6.1, 3.66, "현재 Capa. 대비", w=3.4)
tiles4 = [("현재 연간 Capa.", "26.2", "톤", "제공 기준값", RED, "FDE9E7"),
          ("2026 제시 출하", t2(K("2026 제시 출하")), "톤", "하이닉스+이지켐", DARK, "F2F2F2"),
          ("출하 / Capa.", f"{K('2026 제시 출하')/26200*100:.1f}", "%", "설비 가동률 아님", DARK, "FFFFFF"),
          ("생산 가능량 (추정)", t2(K("2026 생산 가능량 합계 (선택 기준)")), "톤", f"{K('2026 생산 가능 Batch 합계'):.0f} Batch · 달력 연속", DARK, "FFFFFF")]
for k, (h_, v_, u_, s_, c_, f_) in enumerate(tiles4):
    x = 6.1 + (k % 2) * 1.73; y = 3.9 + (k // 2) * 0.72
    add_text(s4, x, y, 1.65, 0.65, [[(h_, {"bold": True, "size": 6.8})], [(v_, {"size": 15, "bold": True, "color": c_}), (" " + u_, {"size": 7, "color": c_})],
                                    [(s_, {"size": 5.8, "color": GRAY})]], align="c", anchor="m", fill=f_, line="BFBFBF", shape=MSO_SHAPE.RECTANGLE, margin=0.03)
# findings
q4 = sum(M[f"AH{m0+k}"].value for k in (9, 10, 11))
add_text(s4, 0.45, 6.08, 9.1, 0.63,
         [[("① ", {"bold": True, "color": RED, "size": 6.8}), (f"9~12월 월 출하 1.94~2.04톤 = 10.2~10.7 Batch 상당 ≈ 월 생산 가능량(10~11 Batch) → 월 여유 −0.14~+0.15톤, 재고 활용 필요", {"size": 6.8})],
          [("② ", {"bold": True, "color": RED, "size": 6.8}), ("원자료 Total 8 Batch(1,520 kg) 기준이면 9월 출하 1,940 kg 대비 부족 → Total 집계 대상(완료/양품/출하) 확인 · 번호 개수와 구분", {"size": 6.8})],
          [("③ ", {"bold": True, "color": RED, "size": 6.8}), (f"10~12월 생산 가능 − 출하 합계 {q4:+,.0f} kg → 2027년 공사 대비 선행 재고는 현재고·기초재고에 의존 (재고 미확인) · 연 생산 가능 − 출하 {M[f'AI{m0+11}'].value/1000:+.2f}톤은 CXMT 미포함·확정 판매 가능량 아님", {"size": 6.8})]],
         fill="F2F2F2", shape=MSO_SHAPE.RECTANGLE, margin=0.06, anchor="m", spacing=1)
notes(s4, f"""[4장 2026년 월별 생산계획 · 출하 비교]
■ 출하(단위 kg, 원자료 실적/계획 구분 미확인): 하이닉스 1,600/1,420/1,440/1,580/1,200/1,200/1,360/1,420/1,380/1,480/1,480/1,420 = 16,980. 이지켐 9~12월 560(4용기/월) = 2,240. 제시 합계 19,220 kg(19.22톤). CXMT 2026 물량은 미제시 → 합계 미포함.
■ 현재 Capa. 26.2톤 대비 19.22톤 = {K('2026 제시 출하')/26200*100:.1f}%, 산술 차이 6.98톤 (가동률·추가 판매 가능량 아님).
■ 과거 생산계획(6~10월): 번호×190 = 2,090/2,090/2,090/1,900/2,090 (10,260 kg), Total×190 = 1,520/2,090/1,520/1,520/2,090 (8,740 kg). 어느 쪽도 확정 생산량 아님. 1~4·11~12월 계획 자료 없음, 5월은 일부.
■ 생산 가능량(추정): 03 시트 선택 기준(기본 = 과거 일정 관측 {K('현재 조건 적용 간격'):.2f} h/Batch), 월 가용시간 = 달력일×24 h(정지·보수 미입력 → 미반영), 월말 진행 Batch 이월 반영, 190 kg/Batch, 양품률 미입력 → 100% 상한. 2026년 {K('2026 생산 가능 Batch 합계'):.0f} Batch = {f0(K('2026 생산 가능량 합계 (선택 기준)'))} kg.
■ 월별 생산 가능 − 출하(kg): {', '.join(diff)} → 누적 {M[f'AI{m0+11}'].value:+,.0f} kg(기초재고 제외).
■ 5 Gal 충진: 2026 월 최대 {K('2026 5 Gal 월 최대 병수'):.0f}병 = {K('2026 5 Gal 월 최대 작업'):.0f} h/월 (1교대 176 h 참고 시 {K('2026 5 Gal 부하율 1교대')*100:.0f}%). 200 L: 9~12월 4용기 × 8 h = 32 h/월.
■ 재고: 기말재고 = 기초재고 + 출하 가능 양품 생산량 − 출하량. 기초재고·목표재고·합격률 미입력 → Excel 04 시트에서 재고 계산 보류 (입력 시 자동 계산, 하이닉스 우선 배정).
■ 확인: 하이닉스 월별 실적/계획 구분, CXMT 2026 물량, Total 집계 기준, 현재고.""")

# ------------------------------------------------------------------ slide 5
clean_slide(s5, "2027년 월별 출하계획과 공급 대응",
            "출하계획 38.02톤 vs 2027년 실제 생산 가능량 · 공급 대응", 5)
section(s5, 0.45, 1.08, "2027년 월별 출하 · 생산 가능량 · 필요 Batch", w=8, sub="단위 kg · 하이닉스 = 22,720 ÷ 12 (표시 반올림)")
k0 = m0 + 12
hx7 = [f1(M[f"T{k0+k}"].value) for k in range(12)]
cx7 = [f0(M[f"U{k0+k}"].value) for k in range(12)]
ez7 = [f0(M[f"V{k0+k}"].value) for k in range(12)]
tt7 = [f1(M[f"W{k0+k}"].value) for k in range(12)]
nb7 = [f1(M[f"X{k0+k}"].value) for k in range(12)]
sr27 = MAP["SR"]["2027"]
pc7 = [f0(M[f"S{k0}"].value)] + ["미확인"] * 11
gb7 = [f"{M[f'Y{k0+k}'].value}/{M[f'Z{k0+k}'].value:.0f}" for k in range(12)]
a7 = [f"{M[f'AA{k0+k}'].value}/{M[f'AD{k0+k}'].value:.0f}" for k in range(12)]
rows5 = [hdr,
         ["운전 조건", "현재", "리플럭스 공사", "", "", "", "시운전", "리플럭스 적용 · 초기 안정화", "", "", "", "", "", ""],
         ["■ SK하이닉스"] + hx7 + [f0(M[f"T{sr27}"].value)],
         ["■ CXMT"] + cx7 + [f0(M[f"U{sr27}"].value)],
         ["■ 이지켐"] + ez7 + [f0(M[f"V{sr27}"].value)],
         ["출하 합계"] + tt7 + [f0(M[f"W{sr27}"].value)],
         ["필요 Batch (÷190)"] + nb7 + [f1(M[f"X{sr27}"].value)],
         ["생산 가능량 (추정)", pc7[0], "공사 중 가용시간 미확인", "", "", "", "양품 미확인", "리플럭스 정제시간·Batch량 미확인", "", "", "", "", "", f"부분 {f0(M[f'S{k0}'].value)}"],
         ["5 Gal 병 / 충진 h"] + gb7 + [f"{M[f'Y{sr27}'].value}/{M[f'Z{sr27}'].value:.0f}"],
         ["200 L 용기 / 충진 h"] + a7 + [f"{M[f'AA{sr27}'].value}/{M[f'AD{sr27}'].value:.0f}"]]
fills = [None, ["F2F2F2", "FFFFFF", "FFE9D6", None, None, None, "FFF2CC", "FDE9E7", None, None, None, None, None, "F2F2F2"],
         None, None, None, ["F2F2F2"] * 14, None, [None, "FFFFFF", UNK, None, None, None, UNK, UNK, None, None, None, None, None, "F2F2F2"], None, None]
colors = [None, None, [RED] + [None] * 13, [CX] + [None] * 13, [ORG] + [None] * 13, None, None, [None, None, "C55A11", None, None, None, "C55A11", "C55A11"] + [None] * 6, None, None]
bolds = [None, [True] * 14, None, None, None, [True] * 14, None, [True] * 14, None, None]
table(s5, 0.45, 1.32, [1.1] + [0.565] * 12 + [0.82], 0.205, rows5, size=6.2, fills=fills, colors=colors, bolds=bolds,
      merges=[(1, 2, 1, 5), (1, 7, 1, 12), (7, 2, 7, 5), (7, 7, 7, 12)])
add_text(s5, 0.45, 3.6, 9.1, 0.16, [[("※ 병수·용기 수는 고객별 20 kg·140 kg 단위 올림 · 200 L 충진 h는 ARS 시간 미확인으로 수동 8 h/용기 참고 · 재고는 기초재고 미입력으로 보류 (1월 생산 가능 − 출하 −643 kg) · 동일 Batch 다고객 배분 시 정제시간 중복 계산 안 함", {"size": 6, "color": GRAY})]], margin=0)
# scenarios
section(s5, 0.45, 3.8, "공급 대응 시나리오 (참고 · 확정 아님)", w=5.2, sub="필요 선행재고 = 누적 (생산 − 출하) 최저점")
sc_rows = [["시나리오 가정", "2027 생산", "연간 생산 − 출하", "필요 선행재고", "최저 월"],
           ["S1 공사 중 중단 · 7월 즉시 47.2 환산", f"{t2(K('S1 2027 생산'))} 톤", f"{K('S1 연간 차이')/1000:+.2f} 톤", f"{t2(K('S1 필요 선행재고'))} 톤", K("S1 최대 부족 월")],
           ["S2 공사 중 현재 조건 유지 · 7월 즉시 47.2", f"{t2(K('S2 2027 생산'))} 톤", f"{K('S2 연간 차이')/1000:+.2f} 톤", f"{t2(K('S2 필요 선행재고'))} 톤", K("S2 최대 부족 월")],
           ["S3 공사 중 50% · 안정화 3개월(현재 수준)", f"{t2(K('S3 2027 생산'))} 톤", f"{K('S3 연간 차이')/1000:+.2f} 톤", f"{t2(K('S3 필요 선행재고'))} 톤", K("S3 최대 부족 월")]]
table(s5, 0.45, 4.03, [2.2, 0.72, 0.95, 0.85, 0.5], 0.235, sc_rows, size=6.3,
      colors=[None, [None, None, RED, RED, None], [None, None, RED, RED, None], [None, None, RED, RED, None]],
      bolds=[None, [False, True, True, True, False], [False, True, True, True, False], [False, True, True, True, False]])
add_text(s5, 0.45, 5.0, 5.22, 0.45, [[("공통: 1월 현재 조건 생산 가능량, 시운전 양품 0, 리플럭스 안정 후 47.2톤 ÷ 12 = 3,933 kg/월(상한 참고), 양품률 100% · ", {"size": 6, "color": GRAY}),
                                     ("최선(S2)에도 연 4.2톤 부족 → 선행재고 7.9톤 이상 필요", {"size": 6, "bold": True, "color": RED})],
                                    [(f"7~12월 출하 대응에 필요한 실효 간격 ≤ {K('7~12월 필요 간격'):.1f} h/Batch (월 {K('7~12월 월 평균 필요 Batch'):.1f} Batch, 190 kg) — 현재 {K('관측 추정 간격'):.1f} h", {"size": 6, "bold": True})]], margin=0)
# resources
section(s5, 5.85, 3.8, "정제 · 충진 자원 제약", w=3.6)
res = [("정제기", f"7~12월 필요 {K('7~12월 월 평균 필요 Batch'):.1f} Batch/월 ↔ 47.2톤 환산 {K('47.2 Batch 상당/월'):.1f} Batch/월 · 리플럭스 정제시간 미확인"),
       ("글로브 박스", f"4월~ 5 Gal {K('2027 5 Gal 월 병수 (4월~)'):.0f}병 = {K('2027 5 Gal 월 최대 작업'):.0f} h/월 → 1교대 176 h의 {K('2027 5 Gal 부하율 1교대')*100:.0f}% · 2교대 {K('2027 5 Gal 부하율 2교대')*100:.0f}%"),
       ("200 L ARS", f"이지켐 {K('2027 200 L 월 용기'):.0f}용기/월 = {K('2027 200 L 월 최대 작업'):.0f} h/월 (수동 8 h 참고) · ARS 시간 미확인, Capa. 미가산"),
       ("Product Tank", f"47.2톤 등가 간격 {K('47.2톤 등가 간격'):.1f} h 시 1교대 충진(경과 48 h)은 Tank 회전 {K('Tank 1교대 여유 vs 47.2'):+.1f} h → 2교대({K('Tank 2교대 여유 vs 47.2'):+.1f} h) 또는 Tank 추가 검토"),
       ("검사", "FQC·OQC 2 h/Batch — 부하 작음 · 검사 대기·장비 공유 확인")]
paras = [[(a + "  ", {"bold": True, "size": 6.6, "color": RED}), (b_, {"size": 6.6})] for a, b_ in res]
add_text(s5, 5.85, 4.03, 3.7, 1.4, paras, fill="FFFFFF", line="BFBFBF", shape=MSO_SHAPE.RECTANGLE, margin=0.06, spacing=2.5, anchor="m")
# actions
section(s5, 0.45, 5.5, "대응 방향 · 추가 확인사항", w=8)
acts = [("선행 생산·재고", "2026년 말 현재고 확인 → 필요 선행재고(7.9~16.0톤) 대비 부족분 산정 · 하이닉스 출하·목표재고 우선 배정 후 CXMT·이지켐"),
        ("공사 중 가동", "2~5월 정제기 가동 가능 여부 확정 (S1↔S2 부족 차이 약 8톤) · 시운전 양품 출하 가능 여부"),
        ("리플럭스 조건", "실제 정제시간·Batch량·안정화 기간 → 7~12월 필요 간격 42.2 h 이하 달성 여부 확인"),
        ("충진 자원", "글로브 박스 교대(2교대 이상)·Product Tank 수·ARS 운영 개시일·충진시간 확정"),
        ("우선순위", "공급 부족 시 CXMT·이지켐 물량 배정 우선순위 및 출하 조정 협의 (영업·생산관리)")]
paras = [[(f"{i+1}. {a}  ", {"bold": True, "size": 7}), (b_, {"size": 7})] for i, (a, b_) in enumerate(acts)]
add_text(s5, 0.45, 5.73, 9.1, 0.98, paras, fill="F2F2F2", shape=MSO_SHAPE.RECTANGLE, margin=0.07, anchor="m", spacing=1)
notes(s5, f"""[5장 2027년 월별 출하계획과 공급 대응]
■ 출하계획(kg): 하이닉스 22,720 ÷ 12 = 1,893.33/월(표시 1,893.3, 연간 22,720 유지), CXMT 4~12월 580/월 = 5,220, 이지켐 1~12월 840/월 = 10,080 → 합계 38,020 kg = 38.02톤. (기존 '이지켐 1월 한정', '부분합계 28.78톤', 'CXMT 연 7톤'은 폐기) 상반기 {f0(K('2027 H1 출하'))} kg, 하반기 {f0(K('2027 H2 출하'))} kg.
■ 필요 Batch(출하 ÷ 190 kg): 1~3월 14.4, 4~12월 17.4 Batch/월, 연 {M[f'X{sr27}'].value:.1f} Batch.
■ 2027년 실제 생산 가능량: 1월 현재 조건 {K('2027 1월 생산 가능 Batch'):.0f} Batch = {f0(K('2027 1월 생산 가능량'))} kg. 2~5월 공사 중 정제기 가용시간, 6월 시운전 양품, 7월 이후 리플럭스 정제시간·Batch량이 미확인 → 연간 실제 생산량은 확정하지 않음(부분합계). 47.2톤 연간 환산값을 2027년 생산량으로 사용하지 않음.
■ 시나리오(참고, Excel 05): S1 공사 중 중단 → 생산 {t2(K('S1 2027 생산'))}톤, 연 {K('S1 연간 차이')/1000:+.2f}톤, 필요 선행재고 {t2(K('S1 필요 선행재고'))}톤({K('S1 최대 부족 월')} 최저). S2 공사 중 현재 조건 유지 → {t2(K('S2 2027 생산'))}톤, {K('S2 연간 차이')/1000:+.2f}톤, 선행재고 {t2(K('S2 필요 선행재고'))}톤. S3 공사 중 50%·안정화 3개월 → {t2(K('S3 2027 생산'))}톤, {K('S3 연간 차이')/1000:+.2f}톤, 선행재고 {t2(K('S3 필요 선행재고'))}톤. 리플럭스 안정 후 3,933 kg/월은 47.2톤 균등 환산 상한 참고.
■ 7~12월 출하 19,880 kg = 104.6 Batch → 184일(4,416 h) 기준 필요 실효 간격 {K('7~12월 필요 간격'):.2f} h/Batch 이하 (현재 관측 {K('관측 추정 간격'):.2f} h).
■ 자원: 5 Gal 4월~ 월 124병(하이닉스 95 + CXMT 29) × 2 h = 248 h/월 → 1교대(8 h×22일=176 h, 참고 가정) {K('2027 5 Gal 부하율 1교대')*100:.0f}%, 2교대 {K('2027 5 Gal 부하율 2교대')*100:.0f}%. 200 L 6용기 × 8 h = 48 h/월(ARS 시간 미확인 시 수동 기준). Product Tank 1기 가정 시 리플럭스 후 간격이 {K('47.2톤 등가 간격'):.1f} h 수준으로 단축되면 1교대 충진(16 h 작업 = 경과 48 h)으로는 Tank 회전 부족.
■ 재고: 기말재고 = 기초재고 + 출하 가능 양품 생산량 − 출하량 / 필요 생산량 = 출하량 + 목표 기말재고 − 기초재고. 하이닉스 출하·목표재고를 먼저 배정한 뒤 CXMT·이지켐 배정(Excel 04). 기초재고·목표재고·합격률 미입력 → 보류.
■ 확인: 공사 중 가동 여부, 시운전·안정화, 리플럭스 정제시간·Batch량, 현재고, 글로브 박스 교대·Tank 수, ARS 운영 개시·충진시간, 고객별 합격률(이지켐 색도).""")

prs.save(OUT)
print("saved", OUT)
