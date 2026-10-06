# -*- coding: utf-8 -*-
"""SKTC 중국법인 설립 검토 PPT — SK trichem 양식(base.pptx) 적용.

사용: python build_cn_pptx.py <template.pptx> <calc.xlsx> <geojson> <out.pptx>
숫자(비용·점수·일정·거리)는 LibreOffice로 재계산한 calc.xlsx에서 읽는다.
"""
import copy
import datetime as dt
import json
import math
import sys

import openpyxl
from pptx import Presentation
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from pptx_helpers import E, RGBColor, delete, etree, fill, notes, pos, runfmt, set_lines, shapes

import cn_data as D

TPL, CALC, GEO, OUT = sys.argv[1:5]
prs = Presentation(TPL)
SRC_SLIDES = list(prs.slides)
s4 = SRC_SLIDES[3]
CHROME = {"TextBox 2", "직사각형 3", "TextBox 5", "타원 6", "그림 7", "직사각형 160", "직선 연결선[R] 162", "TextBox 10"}
DARK, GRAY, RED, ORG, BLU, GRN, PUR, REDC = "1A1A1A", "7F7F7F", "EB002C", "FF7900", "2E75B6", "548235", "7030A0", "C00000"
ST = {"추천": ("E2EFD9", "385723"), "조건부": ("FFF2CC", "7F6000"), "확인 필요": ("FDE9E7", "C00000"), "제외": ("D9D9D9", "404040"),
      "근거 확인": ("E2EFD9", "385723"), "부분 맞음": ("FFF2CC", "7F6000"), "근거 없음": ("FDE9E7", "C00000"), "가정": ("FCE4D6", "843C0C"),
      "보류": ("F2F2F2", "595959"), "점수 보류": ("F2F2F2", "595959")}

# ------------------------------------------------------------------ Excel 값
wb = openpyxl.load_workbook(CALC, data_only=True)
cm = wb["15_CostModel"]
COST = {cm.cell(r, 2).value: (cm.cell(r, 3).value, cm.cell(r, 4).value, cm.cell(r, 5).value, cm.cell(r, 6).value) for r in range(6, 11)}
PERS = [(cm.cell(r, 1).value, cm.cell(r, 2).value, cm.cell(r, 3).value, cm.cell(r, 4).value, cm.cell(r, 5).value) for r in range(14, 17)]
CAPV = (cm["C19"].value, cm["F19"].value)
cc = wb["08_CityComparison"]
SCORE = {cc.cell(r, 1).value: (cc.cell(r, 17).value, cc.cell(r, 18).value, cc.cell(r, 19).value) for r in range(13, 13 + len(D.CITIES))}
dl = wb["12_LogisticsStaff"]
DIST = {dl.cell(r, 1).value: (dl.cell(r, 5).value, dl.cell(r, 6).value, dl.cell(r, 7).value) for r in range(6, 6 + len(D.CITIES))}
tl = wb["16_Timeline"]
TASK = {tl.cell(r, 1).value: dict(name=tl.cell(r, 2).value, s=tl.cell(r, 11).value, e=tl.cell(r, 12).value, eo=tl.cell(r, 14).value,
                                  ep=tl.cell(r, 15).value, so=tl.cell(r, 19).value, sp=tl.cell(r, 20).value) for r in range(6, 6 + len(D.TASKS))}
FX = wb["13_FXAssumptions"]["C6"].value


def f0(v):
    return f"{v:,.0f}" if isinstance(v, (int, float)) else str(v)


def d8(v):
    return f"{v:%Y-%m-%d}" if isinstance(v, (dt.date, dt.datetime)) else str(v)


def md(v):
    return f"{v.year % 100}.{v.month}.{v.day}" if isinstance(v, (dt.date, dt.datetime)) else str(v)


# ------------------------------------------------------------------ 슬라이드 공통
def dup_slide(src):
    new = prs.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    rmap = {}
    for rid, rel in src.part.rels.items():
        if rel.reltype.endswith("/notesSlide") or rel.reltype.endswith("/slideLayout"):
            continue
        rmap[rid] = (new.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref) if rel.is_external
                     else new.part.relate_to(rel.target_part, rel.reltype))
    for el in src.shapes._spTree:
        if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        c = copy.deepcopy(el)
        for node in c.iter():
            for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                if node.get(attr) in rmap:
                    node.set(attr, rmap[node.get(attr)])
        new.shapes._spTree.append(c)
    for sh in list(new.shapes):
        if sh.name not in CHROME:
            delete(sh)
    return new


SLIDES = []


def slide(title, sub):
    s = dup_slide(s4)
    SLIDES.append(s)
    num = len(SLIDES)
    d = shapes(s)
    set_lines(d["TextBox 2"], [title]); set_lines(d["TextBox 5"], [sub])
    for r_ in d["TextBox 5"].text_frame.paragraphs[0].runs:
        r_.font.size = Pt(15 if len(sub) <= 40 else 13.5)
    set_lines(d["타원 6"], [str(num)]); set_lines(d["TextBox 10"], [f"- {num} -"])
    if num >= 10:
        tf = d["타원 6"].text_frame; tf.margin_left = tf.margin_right = 0
        for r_ in tf.paragraphs[0].runs:
            r_.font.size = Pt(7.5)
    return s


def tb(s, x, y, w, h, paras, align="l", anchor=MSO_ANCHOR.TOP, wrap=True):
    """paras: [ [(text, size, bold, color), ...], ... ] 또는 단일 문자열"""
    sh = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = sh.text_frame; tf.word_wrap = wrap; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.03); tf.margin_top = tf.margin_bottom = Inches(0.02)
    if isinstance(paras, str):
        paras = [[(paras, 12, False, DARK)]]
    for i, pr in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        for t, size, bold, color in pr:
            r = p.add_run(); r.text = t; runfmt(r, size, bold, color)
    return sh


def rect(s, x, y, w, h, fc, lc=None, shape=MSO_SHAPE.RECTANGLE, dash=False, lw=0.75):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    if fc:
        fill(sh, fc)
    else:
        sh.fill.background()
    if lc:
        sh.line.color.rgb = RGBColor.from_string(lc); sh.line.width = Pt(lw)
        if dash:
            sh.line.dash_style = MSO_LINE.DASH
    else:
        sh.line.fill.background()
    return sh


def textin(sh, paras, align="c", anchor=MSO_ANCHOR.MIDDLE, ml=0.05):
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(ml); tf.margin_top = tf.margin_bottom = Inches(0.02)
    for i, pr in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER}[align]
        for t, size, bold, color in pr:
            r = p.add_run(); r.text = t; runfmt(r, size, bold, color)


def sec(s, y, main, sub="", x=0.31, w=9.37):
    rect(s, x, y + 0.04, 0.06, 0.2, RED)
    runs = [(main, 12, True, DARK)] + ([("   " + sub, 10, True, REDC)] if sub else [])
    tb(s, x + 0.12, y - 0.01, w - 0.12, 0.3, [runs])


def badge(s, x, y, w, h, label, size=10.5):
    fc, tc = ST.get(label, ("F2F2F2", DARK))
    sh = rect(s, x, y, w, h, fc, tc, MSO_SHAPE.ROUNDED_RECTANGLE, lw=0.75)
    textin(sh, [[(label, size, True, tc)]])
    return sh


def src(s, text):
    tb(s, 0.31, 6.70, 9.37, 0.26, [[("출처: ", 8, True, GRAY), (text, 8, False, GRAY)]])


def _ln(tcPr, tag, spec):
    ln = etree.SubElement(tcPr, qn(tag))
    if spec is None:
        ln.set("w", "0"); etree.SubElement(ln, qn("a:noFill")); return
    col, w = spec; ln.set("w", str(w))
    sf = etree.SubElement(ln, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr")).set("val", col)


def table(s, x, y, colw, rows, fs=12, rh=None, fills=None, bolds=(), colors=None, left_cols=None, hdr_rows=1, status_col=None, wrap=True):
    fills, colors = fills or {}, colors or {}
    nr, nc = len(rows), len(colw)
    rh = rh or [0.34] * nr
    if not isinstance(rh, list):
        rh = [rh] * nr
    gf = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(colw)), Inches(sum(rh)))
    tbl = gf.table
    tp = tbl._tbl.tblPr; tp.set("firstRow", "0"); tp.set("bandRow", "0")
    for el in tp.findall(qn("a:tableStyleId")):
        tp.remove(el)
    for j, w in enumerate(colw):
        tbl.columns[j].width = Emu(int(w * E))
    for i in range(nr):
        tbl.rows[i].height = Emu(int(rh[i] * E))
    left_cols = range(nc) if left_cols is None else left_cols
    for i, row in enumerate(rows):
        hdr = i < hdr_rows
        for j in range(nc):
            v = row[j] if j < len(row) else ""
            c = tbl.cell(i, j)
            tf = c.text_frame; tf.word_wrap = wrap
            p = tf.paragraphs[0]
            lines = str(v if v is not None else "").split("\n")
            for k, ln in enumerate(lines):
                pp = p if k == 0 else tf.add_paragraph()
                r = pp.add_run(); r.text = ln if ln else " "
                col = colors.get((i, j), "FFFFFF" if hdr else DARK)
                bold = hdr or (i in bolds)
                if status_col is not None and j == status_col and not hdr and ln in ST:
                    col = ST[ln][1]; bold = True
                if not hdr and ln.startswith("−"):
                    col = REDC
                runfmt(r, fs, bold, col)
                pp.alignment = PP_ALIGN.LEFT if (j in left_cols and not hdr) else PP_ALIGN.CENTER
            c.margin_left = c.margin_right = Emu(45720); c.margin_top = c.margin_bottom = Emu(18000)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tcPr = c._tc.get_or_add_tcPr()
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB", "a:solidFill", "a:noFill"):
                for el in tcPr.findall(qn(tag)):
                    tcPr.remove(el)
            _ln(tcPr, "a:lnL", None); _ln(tcPr, "a:lnR", None)
            _ln(tcPr, "a:lnT", ("404040", 12700) if i == 0 else None)
            _ln(tcPr, "a:lnB", ("404040", 9525) if (hdr and i == hdr_rows - 1) or i == nr - 1 else ("D9D9D9", 6350))
            bg = fills.get((i, j), fills.get((i, -1), "595959" if hdr else "FFFFFF"))
            if status_col is not None and j == status_col and not hdr and str(v) in ST:
                bg = ST[str(v)][0]
            sf = etree.SubElement(tcPr, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr")).set("val", bg)
    return tbl


def bullets(s, x, y, w, h, items, fs=12.5, gap=4):
    sh = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = sh.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.03)
    for i, it in enumerate(items):
        lvl = 0
        if isinstance(it, tuple):
            it, lvl = it
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        r = p.add_run(); r.text = ("• " if lvl == 0 else "   – ") + it
        runfmt(r, fs if lvl == 0 else fs - 1.5, False, DARK if lvl == 0 else "404040")
    return sh


def chev(s, x, y, w, h, steps, fs=10.5):
    n = len(steps); cw = w / n
    for i, (t1, t2, fc, tc) in enumerate(steps):
        sh = s.shapes.add_shape(MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON, Inches(x + i * cw), Inches(y), Inches(cw + 0.06), Inches(h))
        sh.adjustments[0] = 0.25; sh.shadow.inherit = False; fill(sh, fc); sh.line.fill.background()
        tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.05 if i == 0 else 0.2); tf.margin_right = Inches(0.08)
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = t1; runfmt(r, fs, True, tc)
        if t2:
            p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
            r = p2.add_run(); r.text = t2; runfmt(r, fs - 2, False, tc)


# ------------------------------------------------------------------ 지도
GJ = json.load(open(GEO, encoding="utf-8"))
PROV = {f["properties"]["name"]: f["geometry"] for f in GJ["features"]}
BB = (116.8, 122.25, 29.7, 33.05)


def clip(poly, bb):
    x0, x1, y0, y1 = bb
    def cl(pts, inside, inter):
        out = []
        for i in range(len(pts)):
            a, b = pts[i - 1], pts[i]
            ia, ib = inside(a), inside(b)
            if ib:
                if not ia:
                    out.append(inter(a, b))
                out.append(b)
            elif ia:
                out.append(inter(a, b))
        return out
    def ix(xc):
        return lambda a, b: (xc, a[1] + (b[1] - a[1]) * (xc - a[0]) / (b[0] - a[0]))
    def iy(yc):
        return lambda a, b: (a[0] + (b[0] - a[0]) * (yc - a[1]) / (b[1] - a[1]), yc)
    pts = poly
    for inside, inter in ((lambda p: p[0] >= x0, ix(x0)), (lambda p: p[0] <= x1, ix(x1)),
                          (lambda p: p[1] >= y0, iy(y0)), (lambda p: p[1] <= y1, iy(y1))):
        if not pts:
            break
        pts = cl(pts, inside, inter)
    return pts


def draw_map(s, X, Y, W, highlight, others, points, customers, legend=True, bb=BB, plabels=None):
    lon0, lon1, lat0, lat1 = bb
    k = math.cos(math.radians((lat0 + lat1) / 2))
    H = W * (lat1 - lat0) / ((lon1 - lon0) * k)
    def P(lon, lat):
        return X + (lon - lon0) / (lon1 - lon0) * W, Y + (lat1 - lat) / (lat1 - lat0) * H
    rect(s, X, Y, W, H, "EAF2F8", "BFBFBF")       # 바다·바깥
    for nm_, geom in PROV.items():
        style = ("FDEDF0", RED, 1.25) if nm_ in highlight else (("F7F7F7", "8C8C8C", 0.75) if nm_ in others else ("FBFBFB", "BFBFBF", 0.5))
        polys = geom["coordinates"] if geom["type"] == "MultiPolygon" else [geom["coordinates"]]
        for poly in polys:
            ring = clip([tuple(p) for p in poly[0]], bb)
            if len(ring) < 3:
                continue
            pts = [P(a, b) for a, b in ring]
            ff = s.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]), scale=1.0)
            ff.add_line_segments([(Inches(px), Inches(py)) for px, py in pts[1:]], close=True)
            sh = ff.convert_to_shape(); sh.shadow.inherit = False
            fill(sh, style[0]); sh.line.color.rgb = RGBColor.from_string(style[1]); sh.line.width = Pt(style[2])
    plabels = plabels or {"江苏 강소성": (119.6, 32.85), "浙江 저장성": (120.2, 29.95), "安徽 안후이성": (117.35, 32.75), "上海 상하이": (121.75, 30.98)}
    for nm_, (lon, lat) in plabels.items():
        px, py = P(lon, lat)
        tb(s, px - 0.55, py - 0.1, 1.1, 0.22, [[(nm_, 8.5, True, "8C8C8C" if "강소" not in nm_ else RED)]], align="c")
    for lab, lon, lat, col, dx, dy, bold in points:
        px, py = P(lon, lat)
        rect(s, px - 0.06, py - 0.06, 0.12, 0.12, col, "FFFFFF", MSO_SHAPE.OVAL, lw=0.5)
        if dx < 0:
            tb(s, px - 1.82, py + dy, 1.72, 0.22, [[(lab, 8.5, bold, DARK)]], align="r")
        else:
            tb(s, px + dx, py + dy, 1.72, 0.22, [[(lab, 8.5, bold, DARK)]])
    for lab, lon, lat, dx, dy in customers:
        px, py = P(lon, lat)
        rect(s, px - 0.08, py - 0.08, 0.16, 0.16, "1F3864", "FFFFFF", MSO_SHAPE.DIAMOND, lw=0.5)
        if lab:
            tb(s, px + dx, py + dy, 1.8, 0.22, [[(lab, 8.5, True, "1F3864")]])
    if legend:
        tb(s, X, Y + H + 0.02, W, 0.25, [[("● 후보  ◆ 고객(공개 주소·구 단위)  — 성 경계: 공개 GeoJSON 단순화 · 위치는 대략", 8, False, "595959")]])
    return H


# ================================================================== 01 표지
s = dup_slide(s4); SLIDES.append(s)
d = shapes(s)
for nm_ in ("TextBox 5", "타원 6", "TextBox 10"):
    delete(d[nm_])
set_lines(d["TextBox 2"], ["SKTC 중국법인 설립 검토"])
pos(d["TextBox 2"], x=0.8, y=2.35, w=8.4, h=0.7)
for r_ in d["TextBox 2"].text_frame.paragraphs[0].runs:
    r_.font.size = Pt(34)
tb(s, 0.8, 3.25, 8.4, 0.5, [[("강소성 우선 입지 · 인허가 · 비용 · 실행계획", 20, True, DARK)]])
rect(s, 0.8, 3.9, 2.2, 0.05, RED)
tb(s, 0.8, 4.1, 8.4, 1.2, [[("기준일 2026-10-06 (내부 회의일) · 해외영업/마케팅 1차 검토안", 13, False, "404040")],
                           [("법무·기획·인사·현지 검증 전 자료 — 결론은 조건부이며 기관 확인 결과에 따라 바뀔 수 있음", 12, False, REDC)],
                           [("원자료·출처·수식: SKTC_중국법인설립_Raw_File_Data_20261006.xlsx", 11, False, GRAY)]])
notes(s, "표지. 이 자료는 2026-10-06 내부 회의 지시에 따라 해외영업/마케팅이 만든 1차 검토안입니다. 중국 정부 사이트 원문은 작업 환경 제약으로 직접 열지 못했고, 검색 결과 요약으로 확인한 내용은 '공식 출처·요약 확인'으로 표시했습니다. 법무·현지 전문가 검토 전이므로 결론은 조건부입니다.")

# ================================================================== 02 경영진 요약
s = slide("무석 법인 + 무저장 허가 + 외부 창고가 1순위입니다", "단, 연말 기한 대상 여부·창고 품목 확인 전까지 최종 확정은 불가합니다")
rows = [["핵심 질문", "현재 답 (2026-10-06)", "상태"],
        ["① TC 독자 법인 필요?", "개정안이 확정되면 등록은 '중국 내 수입기업' 명의만 가능 → 등록·자료를 TC가 쥐려면 필요", "조건부"],
        ["② 강소성 어디?", "무석 신오구 법인(무저장 경영허가) + 무석 관할 인정 화공원구 창고 / 차선: 쑤저우+장자강", "조건부"],
        ["③ 강소성이 막히면?", "막힐 곳은 '품목을 받아줄 창고' → 상하이 항만 물류·허페이로 전환 (국가 규정은 동일)", "확인 필요"],
        ["④ 수입만 해도 생산설비?", "생산허가는 생산기업 대상. 수입·판매는 경영허가로 충분 (충진·소분 시 별도)", "근거 확인"],
        ["⑤ 등록은 누구 명의?", "현행: TC 본사+중국 대리인 가능 / 개정 초안: 중국 생산·수입기업만", "확인 필요"],
        ["⑥ 최소 요건·예산?", f"4~5명(전담 안전관리자 포함)·사무실·3PL / 현지 인건비 연 {PERS[0][3] / 10000:,.0f}만 CNY(가정), 그 외 견적 필요", "가정"],
        ["⑦ 연말 못 맞추면?", "기존 备案 보유자가 12-31 전 전환 신청 · 적격 수입자 경유 공급 · 고객 사전 안내", "확인 필요"]]
table(s, 0.31, 1.08, [2.15, 6.17, 1.05], rows, fs=11.5, rh=[0.32] + [0.6] * 7, status_col=2, left_cols=(0, 1))
rect(s, 0.31, 5.72, 9.37, 0.9, "FFF2E6")
tb(s, 0.45, 5.76, 9.1, 0.85, [[("승인 요청  ", 12, True, REDC), ("① 2주 확인 착수 (备案 보유자·제품 분류·MP 실사·무석 기관 질의)  ② 중국 법무·등록 전문가 위촉  "
                                                         "③ 11월 중순 구조·지역·예산 최종 결정", 12, False, DARK)]])
src(s, "EVD-001~003 생태환경부 12호령·개정 초안·환판고체함〔2026〕301호 / EVD-009·011 위험화학품안전법 / EVD-014 무석 세칙 (Excel 02·04)")
notes(s, f"""핵심 메시지: 지금 추천할 수 있는 구조는 'TC 독자 법인(무석 신오구) + 무저장 위험화학품 경영허가 + 허가된 외부 창고'이며, 확정은 2주 확인 후입니다.
근거: ① 생태환경부 12호령 개정 초안(2026-06-11)은 신청 주체를 중국 내 생산·수입기업으로 한정(EVD-002, 초안). ② 환판고체함〔2026〕301호: 2026-08-15 备案 중단, 기존 备案은 12-31까지 등록증(EVD-003). ③ 위험화학품안전법(2026-05-01 시행): 경영 허가제, 생산기업 예외(EVD-008·009). ④ 무석 세칙: 무저장·임차저장 유형(EVD-014).
예상 질문: '연말까지 법인 설립하면 되나?' → 법인 등기와 등록·허가는 별개. 연말 기한은 기존 备案 보유자의 전환 의무이므로 보유자를 먼저 확인해야 함.
예상 질문: '비용은?' → 현지 인건비만 공개 통계로 산정(연 {PERS[0][3]:,.0f} CNY, 가정). 임대·등록·3PL은 견적 필요.
추가 확인: 개정 12호령 최종 공포 여부, TC 제품의 备案 보유자·번호, 무석 신오구 허가 기한.""")

# ================================================================== 03 회의 지시와 목적
s = slide("독자 법인의 명분은 등록 명의와 자료 통제입니다", "회의 지시: 강소성 우선 · 연말 등록 · 최소 자금 · MP 대비 명분 정리")
sec(s, 1.08, "회의에서 나온 내용", "내부 제공 · 검증 전", w=4.6)
bullets(s, 0.31, 1.42, 4.55, 4.9, [
    "중국 고객(CXMT·SK하이닉스 우시·YMTC)에 반도체 전구체 공급 추진",
    "사장님: 기존 MP 중국법인에 기능 추가해 TC가 이용",
    "대표님: TC 독자 법인으로 자체 운영 선호",
    "우선 소재: Y·Sc·La 전구체 3종 (이후 CpZr 등 확대)",
    "외부·고객 명의 등록 시 기술정보 노출 우려",
    "MP는 위험물 인허가 미보유 언급 (미확인)",
    "목표: 2026년 12월 말 법인·물질등록 완료",
    "초기 4명, 중국 투입 자금 최소화, 창고는 향후 재고 거점"], fs=12)
sec(s, 1.08, "이 자료가 답하는 질문", x=5.05, w=4.63)
bullets(s, 5.05, 1.42, 4.6, 4.9, [
    "독자 법인 필요성 vs MP 활용",
    "강소성 내 도시·단지·운영 방식",
    "강소성이 막히는 이유와 인근 대안",
    "수입만 해도 생산허가·설비가 필요한가",
    "신규화학물질 등록 신청 주체",
    "최소 인력·시설·자본·예산·일정",
    "연말 미달 시 합법적 대안"], fs=12)
rect(s, 0.31, 5.95, 9.37, 0.62, "F2F2F2")
tb(s, 0.45, 6.0, 9.1, 0.55, [[("원칙  ", 11.5, True, REDC), ("강소성은 '우선 조사 대상'이지 정해진 결론이 아님 · MP 활용안의 이점도 함께 비교 · 법인 등기 ≠ 사업 개시", 11.5, False, DARK)]])
src(s, "2026-10-06 내부 회의 (내부 제공)")
notes(s, "회의에서 나온 내용은 모두 '내부 제공·검증 전'입니다. 특히 MP의 허가 미보유, 12월 말 기한, 4명 인력은 사실 확인 대상입니다. 독자 법인의 사업 명분은 '빨리 만든다'가 아니라 '등록 명의·자료 접근·재고·고객 대응을 TC가 통제한다'에 있으며, 반론(MP가 이미 허가를 보유한 경우)은 23장에서 비교합니다.")

# ================================================================== 04 고객·제품
s = slide("제품 정보가 없어 규제 판단은 시나리오로 합니다", "고객 위치는 공개 주소 기준 · CAS·SDS·물량·기존 备案 확인이 첫 과제")
sec(s, 1.08, "① 중국 고객", "거래·공급·등록 상태는 내부 확인")
rows = [["고객", "위치 (공개 정보)", "가까운 후보", "확인할 것"],
        ["CXMT (长鑫存储)", "안후이 허페이 경제기술개발구", "허페이 · 난징", "반입 공장·현 수입자·备案 주체"],
        ["SK하이닉스 우시", "강소 무석 신오구", "무석 · 쑤저우", "공급자 변경 승인 절차"],
        ["YMTC (长江存储)", "후베이 우한 동호고신구", "(지도 밖 서쪽)", "공급 여부 자체 확인"]]
table(s, 0.31, 1.36, [2.2, 2.9, 1.8, 2.47], rows, fs=12, rh=[0.32, 0.42, 0.42, 0.42], left_cols=(0, 1, 2, 3))
sec(s, 3.15, "② 대상 제품과 비어 있는 입력값", "제품 코드·화학명·CAS 미제공 → 임의 기재하지 않음")
rows = [["제품", "알고 있는 것", "비어 있는 입력값 (판단에 미치는 영향)"],
        ["Y 전구체", "원소 기준 소재군", "CAS → 신규물질 여부 / 물성 → 위험화학품·창고"],
        ["Sc 전구체", "원소 기준 소재군", "연간 수입량 → 등록 유형(1t·10t 기준)"],
        ["La 전구체", "원소 기준 소재군", "기존 备案·수입자 → 연말 기한 적용 여부"],
        ["CpZr 등 (후속)", "중국 공급 제품", "고객별 수입 구조 → 공급자 변경 필요 여부"]]
table(s, 0.31, 3.43, [2.0, 2.2, 5.17], rows, fs=12, rh=[0.32] + [0.42] * 4, left_cols=(0, 1, 2))
tb(s, 0.31, 5.65, 9.37, 0.9, [[("원소를 파는 것이 아니라 Y·Sc·La를 포함한 전구체 화합물입니다. ", 11.5, True, DARK),
                               ("같은 원소라도 화합물·용매·농도에 따라 신규물질 여부와 위험 분류가 달라지므로 제품별로 따로 판정합니다.", 11.5, False, DARK)]])
src(s, "EVD-029~031 (CXMT·YMTC 공개 주소, SK하이닉스 우시 보도) / 05_Materials 시트")
notes(s, "고객 위치는 공개 주소 또는 언론 보도 기준이며 실제 반입 공장 위치는 고객 확인이 필요합니다. 제품 정보(CAS·SDS·물량·기존 备案)가 없어서 6장의 시나리오 A~D로 분석합니다. 이 정보가 들어오면 규제 판단이 바로 좁혀집니다 — 2주 과제 1번.")

# ================================================================== 05 가설 검증
s = slide("회의 가설 5개 중 3개는 근거와 다르거나 조건부입니다", "생산설비·한국 시험자료·법인 설립 즉시 수입은 사실과 다름")
rows = [["회의 주장", "확인 결과 (2026-10-06)", "판정", "근거"],
        ["해외 제조자는 등록 못 하고\n중국 법인이 있어야 한다", "현행: 중국 대리인 지정 시 해외 생산자도 신청 가능\n개정 초안: 중국 생산·수입기업만 → 확정 시 맞는 말", "부분 맞음", "EVD-001·002"],
        ["수입만 해도 생산허가·\n생산설비가 필요하다", "생산허가는 '생산기업' 대상. 수입·판매는 경영허가\n(충진·소분 등 가공을 하면 별도 판단)", "근거 없음", "EVD-008·011"],
        ["12월 말까지 등록 안 하면\n영업정지 대상", "기존 备案 보유자는 12-31까지 등록증 취득 의무 (통지)\n'영업정지' 제재 근거는 미확인", "부분 맞음", "EVD-003"],
        ["한국 시험자료를\n그대로 쓸 수 있다", "간이등록도 중국 공시생물 생태독성 시험 포함 필요\n→ 일부 보완 시험 가능성 높음", "근거 없음", "EVD-006"],
        ["법인 설립이 끝나면\n바로 수입 가능", "등기 후 경영허가·위험화학품 등록·신규물질 등록·\n중문 SDS·라벨이 갖춰져야 수입·판매", "근거 없음", "EVD-012·018·019"]]
table(s, 0.31, 1.08, [2.45, 4.6, 1.1, 1.22], rows, fs=11.5, rh=[0.32] + [0.86] * 5, status_col=2, left_cols=(0, 1))
src(s, "SRC-001·004·005 생태환경부, SRC-009 위험화학품안전법(2026-05-01), SRC-040 등록 지침 — 검색 요약 확인·원문 대조 필요")
notes(s, """주장 1: 현행 12호령은 해외 생산·무역기업이 중국 대리인을 지정하면 신청인이 될 수 있다(EVD-001). 2026-06 개정 초안은 중국 내 생산·수입기업만 신청인으로 하는 방향(EVD-002) — 최종 공포 여부 확인 필요.
주장 2: 위험화학품안전법은 경영 허가제를 두고, 안전생산허가는 생산기업 대상(EVD-008·011). 수입·판매 법인에 생산설비를 요구하는 규정은 확인되지 않음. 충진·소분·분장은 지역에 따라 '분장·충진 경영' 또는 생산으로 볼 수 있어 별도 확인.
주장 3: 환판고체함〔2026〕301호 — 2026-08-15 备案 접수 중단, 기존 备案 보유자는 12-31까지 등록증 취득(EVD-003). 미취득 시 제재 수준은 원문·법전 조항 확인 필요.
주장 4: 간이등록 생태독성 자료에 중국 공시생물 시험 포함(EVD-006). 한국 자료 인정 범위는 물질별 확인.
주장 5: 등기는 출발점일 뿐.""")

# ================================================================== 06 규제 시나리오
s = slide("제품이 신규·위험물인지에 따라 해야 할 일이 갈립니다", "CAS 확인 전이므로 4가지 경우를 모두 준비 — 운송 위험물 분류는 별도 판정")
rows = [["경우", "신규화학물질 등록", "위험화학품 경영허가·등록", "창고·운송", "사업 영향"],
        ["A 신규 · 위험", "필요", "필요", "허가 창고·위험물 운송", "가장 무거움 · 등록 명의가 핵심"],
        ["B 신규 · 비위험", "필요", "불필요 (판정 근거 보관)", "물성 맞는 일반 창고", "등록 명의 문제만 남음"],
        ["C 기존 · 위험", "불필요", "필요", "허가 창고·위험물 운송", "법인·허가·창고가 핵심"],
        ["D 기존 · 비위험", "불필요", "불필요", "일반", "MP·일반 무역법인도 가능"]]
table(s, 0.31, 1.08, [1.7, 1.7, 2.2, 1.85, 1.92], rows, fs=12, rh=[0.32] + [0.52] * 4, fills={(1, -1): "FFF2E6"}, left_cols=(0, 4))
sec(s, 3.6, "판정에 필요한 입력값", "05_Materials 시트")
bullets(s, 0.31, 3.92, 9.37, 2.6, [
    "CAS·화학명·조성·용매·농도 → IECSC(중국 기존화학물질 목록) 검색으로 신규 여부",
    "SDS의 GHS 분류·물리적 위험 → 위험화학품 목록 해당·판정 (목록은 공고로 갱신됨, 2026년 3호 공고 5종 추가)",
    "UN번호·운송 분류(자연발화·수분반응 여부) → 창고·운송 수용 가능성",
    "연간 수입량(샘플/양산) → 등록 유형 (1톤 미만 종전 备案 → 간이등록 준용)",
    "기존 备案·등록 보유자와 번호 → 연말 기한 적용 대상인지"], fs=12)
src(s, "EVD-003·004 생태환경부 통지, EVD-032 응급관리부 2026년 3호 공고, 05_Materials")
notes(s, "가장 무거운 경우(A)를 기준으로 준비하되, 제품 정보가 들어오면 경우가 줄어듭니다. 운송상 위험물(UN 분류)과 위험화학품 목록 해당 여부는 서로 다른 판정이므로 둘 다 확인해야 합니다. CAS나 위험등급을 임의로 넣지 않았습니다.")

# ================================================================== 07 주체 구분
s = slide("등록·수입·판매·보관 주체마다 필요한 허가가 다릅니다", "B안 기준: 등록·수입·판매는 TC 법인, 보관은 허가 창고 운영사")
chev(s, 0.31, 1.15, 9.37, 0.78, [("TC 본사 (한국)", "제조·수출", "D9D9D9", DARK), ("수입", "통관 명의자", BLU, "FFFFFF"),
                                  ("판매", "중국 내 거래", ORG, "FFFFFF"), ("보관", "창고", GRN, "FFFFFF"), ("고객", "반입", "7F7F7F", "FFFFFF")], fs=12)
rows = [["단계", "누가 (B안)", "필요한 등록·허가", "근거"],
        ["신규물질 등록", "TC 법인 (개정안) / TC 본사+대리인 (현행)", "간이·상규 등록 또는 备案 전환", "EVD-001~004"],
        ["수입", "TC 법인 (수입 명의)", "위험화학품 등록(첫 수입 전) · 중문 SDS·라벨 · 적합성 성명", "EVD-018·019"],
        ["판매", "TC 법인", "위험화학품 경영허가 (무저장)", "EVD-008·015"],
        ["보관", "허가 3PL", "3PL의 창고(저장) 경영허가·품목 범위", "EVD-013"],
        ["제조 (충진·소분)", "해당 없음 (현 단계)", "안전생산허가 또는 분장·충진 경영 — 필요 시 별도", "EVD-011"]]
table(s, 0.31, 2.15, [1.75, 2.85, 3.67, 1.1], rows, fs=12, rh=[0.32] + [0.55] * 5, left_cols=(0, 1, 2), fills={(5, -1): "F2F2F2"})
tb(s, 0.31, 5.65, 9.37, 0.9, [[("포인트  ", 12, True, REDC), ("수입 명의자 = 위험화학품 등록·통관 책임자, 판매자 = 경영허가 보유자, 보관 = 창고 허가 보유자. "
                                                          "같은 회사일 필요는 없지만 계약서에 책임을 나눠 적어야 합니다.", 12, False, DARK)]])
src(s, "SRC-009 위험화학품안전법, SRC-016 등록관리 53호령, SRC-017 해관총서 2020-129호")
notes(s, "용어: 수입자 = 통관 서류상 들여오는 회사, 등록 신청 주체 = 신규물질 등록증 명의자, 외부 창고 = 허가를 가진 3PL. 회의에서 '수입만 해도 생산설비'라는 말은 이 구분이 섞여서 나온 것으로 보입니다. 제조(충진·소분)는 지금 계획이 확정되지 않았으므로 초기 법인에는 넣지 않습니다.")

# ================================================================== 08 운영 구조 비교
s = slide("최소 투자로 등록 명의를 쥐는 구조는 B안입니다", "MP 허가 보유가 확인되면 A안이 가장 빠를 수 있음 — G안은 연말 대응 수단")
rows = [["구조", "등록 명의", "수입·판매", "창고", "법적 가능성", "비용·기간", "판정"],
        ["A MP 활용", "MP", "MP", "MP/3PL", "MP 허가 보유 시", "보유 시 최소·최단", "확인 필요"],
        ["B TC+무저장+3PL", "TC 법인", "TC 법인", "3PL", "가능 (근거 기반)", "낮음 · 3~6개월", "추천"],
        ["C TC+임차 창고", "TC 법인", "TC 법인", "TC 임차", "가능성 (무석 세칙)", "중간 · 6~9개월", "조건부"],
        ["D TC+자체 창고", "TC 법인", "TC 법인", "TC", "원구 입주 전제", "높음 · 12~24개월", "보류"],
        ["E 거점 분리", "TC 법인", "TC 법인", "타 도시 3PL", "관할 확인 필요", "B와 유사", "조건부"],
        ["F 충진·소분 포함", "TC 법인", "TC 법인", "TC", "현 단계 불필요", "매우 높음", "제외"],
        ["G 대리인·수입자", "대리인/수입자", "적격 수입자", "수입자", "현행 가능·개정 시 수입자", "가장 빠름", "조건부"]]
table(s, 0.31, 1.08, [1.85, 1.2, 1.15, 1.0, 1.75, 1.42, 1.0], rows, fs=11.5, rh=[0.32] + [0.5] * 7, status_col=6, left_cols=(0,),
      fills={(2, -1): "FFF2E6"})
tb(s, 0.31, 5.0, 9.37, 1.55, [[("왜 B안인가  ", 12, True, REDC), ("등록·수입·판매 명의를 TC가 갖고, 위험품 보관은 이미 허가를 가진 창고에 맡겨 초기 시설 투자를 피합니다.", 12, False, DARK)],
                               [("B안의 약점  ", 12, True, REDC), ("창고 운영사의 허가 품목 안에 우리 제품이 들어가야 합니다 — 자연발화·수분반응 물질이면 받아줄 창고가 적을 수 있습니다.", 12, False, DARK)]])
src(s, "10_OperatingModels 시트 · EVD-002, EVD-013~015 (기간은 분석 가정)")
notes(s, "A안은 MP 정보가 없어 판정 보류입니다. MP가 해당 품목 경영허가·창고·등록 수행 능력을 이미 갖고 있다면 속도·비용은 A안이 우위입니다(23장 시나리오 1). G안(대리인·적격 수입자)은 법인 설립과 별개로 연말 기한에 대응하는 수단이지만, 등록증 명의가 대리인·수입자에게 남는 문제가 있습니다. 기간은 분석 가정입니다.")

# ================================================================== 09 법인 설립 요건
s = slide("외투 무역법인은 100% 독자 설립이 가능합니다", "등록자본 법정 최저액 없음 · 5년 내 납입 · 법인 등기 ≠ 화학품 사업 개시")
rows = [["항목", "법정 의무", "실무·확인 사항", "근거"],
        ["지분", "화학품 도매는 네거티브리스트 제한 항목 아님(파악)", "원문 대조 필요", "EVD-024"],
        ["등록자본", "법정 최저액 없음 · 설립 후 5년 내 납입", "은행·임대인이 규모를 보기도 함", "EVD-022"],
        ["기관", "법정대표인 = 이사 또는 경리 · 소규모는 감사 1명/생략", "법인장과 법정대표인은 다를 수 있음", "EVD-023"],
        ["주소", "실제 사무 장소 (임대계약)", "무저장 경영허가 장소 요건과 일치", "EVD-015"],
        ["한국 서류", "아포스티유 (영사 인증 불필요)", "번역·공증 기간 포함", "EVD-025"],
        ["한국 신고", "해외직접투자 사전 신고 (외국환은행)", "송금·증권취득 보고", "EVD-026"],
        ["등기 후", "계좌·외환·세무·해관 등록", "기관 실무 기간 확인", "기관 확인"]]
table(s, 0.31, 1.08, [1.3, 3.75, 3.17, 1.15], rows, fs=11.5, rh=[0.32] + [0.47] * 7, left_cols=(0, 1, 2))
chev(s, 0.31, 4.95, 9.37, 0.72, [("투자 신고", "한국", "D9D9D9", DARK), ("서류 준비", "아포스티유", "D9D9D9", DARK), ("등기", "营业执照", BLU, "FFFFFF"),
                                  ("계좌·세무", "해관·외환", BLU, "FFFFFF"), ("경영허가", "무저장", ORG, "FFFFFF"), ("등록", "위험화학품·신규물질", ORG, "FFFFFF")], fs=11)
tb(s, 0.31, 5.8, 9.37, 0.8, [[("주의  ", 12, True, REDC), ("'법인 설립 4명 필수', '등록자본은 반드시 얼마' 같은 법정 기준은 확인되지 않았습니다. 인력은 경영허가 요건(전담 안전관리자 등)에서 나옵니다.", 12, False, DARK)]])
src(s, "SRC-020 공사법(2024-07-01), SRC-021 네거티브리스트 2024, SRC-022 외교부, SRC-023 외국환거래 신고")
notes(s, "법정 의무와 실무를 구분했습니다. 외국인 주재원 취업허가·체류, 현지 직원 근로계약·사회보험, 실질 지배자 정보 신고 등은 등기 대행사·법무 확인 항목입니다. 기존 연락사무소가 있다면 유지·정리 방안을 별도로 검토합니다(내부 확인).")

# ================================================================== 10 등록·IP
s = slide("신규물질 등록 명의자가 자료와 책임을 갖습니다", "개정안이 확정되면 해외 기업은 자료만 제공 — 중국 수입기업 명의가 필수")
rows = [["구분", "현행 (12호령)", "개정 초안 (2026-06)"],
        ["신청 주체", "중국 생산·수입기업 또는 해외 기업 + 중국 대리인", "중국 내 생산·수입기업만"],
        ["소량 (备案)", "1톤 미만 등 备案 가능", "8-15 접수 중단 · 기존 备案은 12-31까지 등록증"],
        ["기존 등록증", "유효", "유효 · 기재 사항 변경 시 재신청"],
        ["시험자료", "간이등록도 중국 공시생물 생태독성 시험 포함", "확인 필요 (개정 세부)"],
        ["정보 보호", "명칭·CAS·구조 비공개 신청 가능 (기간 제한)", "확인 필요"]]
table(s, 0.31, 1.08, [1.6, 3.9, 3.87], rows, fs=12, rh=[0.32] + [0.52] * 5, left_cols=(0, 1, 2))
sec(s, 4.05, "누가 무엇을 보게 되나", "법인을 세워도 '노출'은 남고 '통제권'이 바뀜")
rows = [["상대", "보게 되는 정보", "통제 수단"],
        ["생태환경부·심사기관", "구조·조성·시험자료 전체", "정보보호 신청"],
        ["중국 시험기관", "시험 물질·결과", "시험 계약 비밀유지"],
        ["대리인·수입자(MP·고객)", "신청서·자료 (명의자면 등록증까지)", "NDA로 부족 — 명의가 핵심"]]
table(s, 0.31, 4.35, [2.6, 3.6, 3.17], rows, fs=12, rh=[0.32] + [0.48] * 3, left_cols=(0, 1, 2))
src(s, "EVD-001~007 · SRC-001·004·005·008·036·040")
notes(s, """독자 법인을 만들어도 정부·시험기관에 대한 정보 제공은 없어지지 않습니다. 달라지는 것은 등록증 명의와 자료 접근을 TC가 통제한다는 점입니다.
대리인·고객 명의로 등록하면 계약 종료 후에도 등록증이 그쪽에 남고, 명의 변경은 '기재 사항 변경 → 재신청'이 될 수 있어(EVD-005) 처음부터 최종 명의자로 신청하는 편이 안전합니다.
NDA로 해결되는 것: 대리인·시험기관의 외부 유출 금지. NDA로 해결 안 되는 것: 등록증 명의·법적 책임·정부 제출 자료의 공개 범위.
확인: 개정 최종본의 신청 주체·정보보호 조항, 한국 시험자료 소유권(공동시험 여부).""")

# ================================================================== 11 위험화학품 허가 경로
s = slide("위험화학품이면 TC 법인은 판매(경영) 허가가 필요합니다", "무저장 허가가 최소안 — 저장하려면 지정 구역(화공원구 등) 창고가 필요")
rows = [["허가·등록", "언제 필요", "핵심 조건", "기간 (확인 수준)"],
        ["경영허가 · 무저장", "판매만, 자체 저장 없음", "사무 장소, 책임자·전담 안전관리자 자격, 규정·응급계획", "20 영업일 사례 (무석 확인)"],
        ["경영허가 · 임차저장", "허가 창고를 빌려 보관", "임차 창고 요건·안전관리 협의", "확인 필요"],
        ["저장 경영 (자체)", "자체 창고 보관", "지정 구역 + 안전평가 + 건설 심사", "1년 이상 (가정)"],
        ["위험화학품 등록", "목록 물질 첫 수입 전", "SDS·라벨·분류 보고", "확인 필요"],
        ["통관 검사", "수입 신고 시", "UN번호·적합성 성명·중문 SDS·라벨", "건별"],
        ["안전생산허가", "생산할 때만", "수입·판매에는 해당 없음", "-"]]
table(s, 0.31, 1.08, [2.1, 2.1, 3.5, 1.67], rows, fs=11.5, rh=[0.32] + [0.55] * 6, left_cols=(0, 1, 2), fills={(1, -1): "FFF2E6", (6, -1): "F2F2F2"})
tb(s, 0.31, 4.85, 9.37, 1.7, [[("무석의 구분  ", 12, True, REDC), ("무석 세칙은 경영 방식을 무저장·유저장·임차저장·주유소로 나눕니다. 무저장은 '서류로만 거래'이며 사무실에 위험물을 둘 수 없습니다.", 12, False, DARK)],
                               [("유효기간  ", 12, True, REDC), ("경영허가 3년, 만료 3개월 전 연장 신청. 위험화학품안전법(2026-05-01) 시행 후 하위 규정 개정 여부는 확인 필요.", 12, False, DARK)]])
src(s, "SRC-009 위험화학품안전법, SRC-010 경영허가 55호령, SRC-011 무석 세칙(2023), SRC-035 처리기한 사례")
notes(s, "무저장 경영은 사무실 주소로 허가를 받고, 실제 물건은 허가 창고에 둡니다. 처리기한 20 영업일은 다른 지역 사례이며 무석 신오구 기한은 질의가 필요합니다. 저장 시설은 지방정부가 정한 전용 구역에 있어야 한다는 요건(EVD-013)이 있어, 강소성에서는 인정 화공원구 안 창고가 현실적인 선택입니다.")

# ================================================================== 12 최소 vs 중장기
s = slide("창고·시설·인력은 최소 조건과 중장기 조건이 다릅니다", "최소(B안): 사무실 + 4~5명 + 3PL / 중장기(C·D안): 전용 창고·창고 인력")
rows = [["항목", "최소 — 인허가 요건 (B안)", "중장기 — 사업 운영 (C·D안)"],
        ["사무실", "무저장 경영 장소 (위험물 보관 금지)", "고객 대응·기술지원 공간 확대"],
        ["창고", "허가 3PL 계약 (품목·물성 수용)", "임차 전용 창고 → 자체 창고 (원구 내)"],
        ["필수 인력", "법정대표인 1 + 전담 안전관리자 1 (자격 시험)", "창고·품질 담당 추가"],
        ["운영 인력", "영업 1 + 경리 1 (외주 가능)", "물류·통관 1 + 기술 대응"],
        ["관리 체계", "안전 규정·응급계획·교육", "창고 안전평가·소방·환경 관리"],
        ["저장 기준", "3PL이 GB 15603-2022 준수", "자체 준수 (혼재 금지·입출고)"]]
table(s, 0.31, 1.08, [1.5, 3.9, 3.97], rows, fs=12, rh=[0.32] + [0.52] * 6, left_cols=(0, 1, 2))
tb(s, 0.31, 4.75, 9.37, 1.8, [[("창고 면적  ", 12, True, REDC), ("법정 최소 면적은 확인되지 않았습니다. 면적 = 평균 재고량 ÷ 용기당 용량 × 용기 점유면적 × 통로·이격 계수 — 물량·용기 정보가 들어오면 계산합니다.", 12, False, DARK)],
                               [("4명 구상  ", 12, True, REDC), ("이동선 총감·신규 채용·최영진·경리 중 1명이 '전담' 안전관리자로 자격을 갖추면 B안 최소 운영은 가능 (법정 최소 인원 아님).", 12, False, DARK)]])
src(s, "EVD-012·013·028 · 12_LogisticsStaff 시트")
notes(s, "인허가에 필요한 최소 조건과 사업을 키울 때 필요한 조건을 나눴습니다. 전담 안전관리자는 '겸직 불가'로 해석될 수 있어 영업·물류와 분리하는 것이 안전합니다. 창고 면적은 물량 정보가 없어 계산식만 제시합니다.")

# ================================================================== 13 강소성 지도
s = slide("강소성 남부에 인정 화공원구와 고객이 모여 있습니다", "1단계 선별: 법인 거점 3곳 · 창고 후보는 원구 지위만 확인 (품목은 문의 예정)")
pts = [("무석 신오구", 120.37, 31.55, RED, -1.05, -0.2, True), ("锡山 원구", 120.47, 31.63, ORG, 0.1, -0.13, False),
       ("江阴临港 원구", 120.15, 31.93, ORG, -0.95, 0.05, False), ("宜兴 원구", 119.82, 31.36, ORG, -0.85, -0.1, False),
       ("쑤저우 SIP", 120.72, 31.32, RED, 0.1, -0.05, True), ("장자강 扬子江 원구", 120.43, 31.97, ORG, 0.1, -0.22, False),
       ("常熟 원구", 120.85, 31.72, ORG, 0.1, -0.1, False), ("창저우 滨江", 119.90, 31.98, ORG, -0.95, -0.25, False),
       ("전장 新区", 119.68, 32.18, "A6A6A6", -0.82, -0.12, False), ("난징 江北", 118.85, 32.26, "A6A6A6", 0.0, 0.08, False),
       ("난퉁 경개구", 120.95, 31.95, "A6A6A6", 0.1, -0.1, False), ("양저우 仪征", 119.18, 32.27, "A6A6A6", 0.1, -0.3, False),
       ("泰兴", 120.02, 32.17, "A6A6A6", 0.1, -0.1, False), ("쿤산 (창고 제외)", 120.98, 31.38, "A6A6A6", 0.1, -0.24, False),
       ("상하이", 121.48, 31.23, BLU, -0.6, 0.05, False)]
cus = [("SK하이닉스 우시", 120.40, 31.50, 0.1, 0.0)]
H13 = draw_map(s, 0.31, 1.08, 5.4, {"江苏省"}, {"浙江省", "安徽省", "上海市"}, pts, cus, bb=(118.6, 122.0, 30.75, 32.55),
               plabels={"江苏 강소성": (119.5, 32.47), "安徽 안후이성": (118.95, 31.15), "浙江 저장성": (120.3, 30.85), "上海": (121.6, 31.0)})
tb(s, 0.36, 1.12 + H13 * 0.42, 1.3, 0.45, [[("← CXMT 허페이", 9, True, "1F3864")], [("(서쪽 약 300 km)", 8, False, "1F3864")]])
rows = [["구분", "1단계 결과"], ["법인 거점", "무석 · 쑤저우 SIP · (난징)"], ["창고 후보\n(인정 원구)", "무석 관할 3 · 쑤저우 관할 2\n창저우 · 전장 · 난징 · 난퉁 3"],
        ["창고 제외", "쿤산 (1차 명단에 원구 없음)"], ["확인 필요", "타이창 (항만·2차 명단)"]]
table(s, 5.85, 1.08, [1.15, 2.68], rows, fs=11, rh=[0.3, 0.6, 0.9, 0.6, 0.6], left_cols=(0, 1))
tb(s, 5.85, 4.25, 3.83, 2.3, [[("읽는 법", 11.5, True, REDC)],
                              [("빨강 = 법인·영업 거점 후보, 주황 = 상위 후보 관할 인정 화공원구, 회색 = 보류, 파랑 = 인근 대안. ", 10.5, False, DARK)],
                              [("원구가 있다는 것 ≠ 우리 제품을 보관할 창고가 있다는 것. 창고·품목은 2주 내 문의합니다.", 10.5, True, DARK)]])
src(s, "EVD-020 강소성 화공원구 인정 1차 명단(2023-06-02), EVD-029·031 고객 위치 · 지도 경계: 공개 GeoJSON(단순화)")
notes(s, "강소성 전체 1차 인정 화공원구 23곳 중 고객(무석·허페이)과 수입항(상하이·장자강) 동선에 있는 남부 후보를 표시했습니다. 쑤저우와 쿤산·타이창·장자강·창수는 같은 쑤저우시 관할(현급시)이며, 무석과 강음·의흥도 같은 관계입니다. 북부(롄윈강·쉬저우·화이안 등) 원구는 고객 거리로 1단계에서 제외했습니다. 지도 위치는 도시·구 중심의 대략 좌표입니다.")

# ================================================================== 14 강소성 비교표
s = slide("강소성에서는 무석이 앞서고 쑤저우가 뒤따릅니다", "점수는 분석 가정 (1~5점) — 정보 부족 후보는 점수를 매기지 않음")
ids = ["CITY-001", "CITY-002", "CITY-003", "CITY-006", "CITY-004", "CITY-005", "CITY-009"]
CN = {c[0]: c for c in D.CITIES}
rows = [["도시 (관할)", "역할", "무저장 경영 경로", "창고 후보", "SK하이닉스\n직선 km", "점수", "판정"]]
for cid in ids:
    c = CN[cid]
    sc = SCORE[cid][0]
    lab = "추천" if cid == "CITY-001" else ("조건부" if cid in ("CITY-002",) else ("제외" if cid == "CITY-009" else "보류"))
    rows.append([f"{c[2]} {c[3].split(' ')[0]}", c[4][:16], c[8].replace("(", "\n("), c[9].replace("(", "\n("), f0(DIST[cid][1]),
                 f"{sc:.0f}" if isinstance(sc, (int, float)) else "보류", lab])
table(s, 0.31, 1.08, [1.75, 1.85, 1.45, 1.75, 0.9, 0.67, 1.0], rows, fs=10.5, rh=[0.45] + [0.66] * 7, status_col=6, left_cols=(0, 1),
      fills={(1, -1): "FFF2E6"})
src(s, "08_CityComparison 시트 (관문·점수·가중치 수식), EVD-014·020·021 · 직선거리 = 대략 좌표 기준, 도로·운송시간 아님")
notes(s, "필수 관문(법인 설립·무저장 경영 경로·적격 창고 후보)을 먼저 보고, 관문을 통과한 후보만 점수를 매겼습니다. 점수는 분석 가정이며 기관 회신이 오면 갱신합니다. 쿤산은 1차 인정 명단에 원구가 없어 창고 거점에서 제외(사무 거점으로는 가능).")

# ================================================================== 15~17 상위 후보 상세 (같은 레이아웃)
DET = [
    ("CITY-001", "강소성 후보 1 · 무석: 고객 옆 법인 + 관할 원구 창고", "SK하이닉스 우시와 같은 신오구 · 무석 세칙에 무저장·임차저장 유형이 명시돼 있음",
     [["권장 운영 구조", "B안 (무저장 경영 + 3PL) → 필요 시 C안 (임차저장)"],
      ["법인 주소", "무석 신오구 업무용 건물 (무저장 경영 장소)"],
      ["창고 후보", "江阴临港化工园区 · 锡山경개구 新材料产业园 · 宜兴 新材料产业园 (인정 원구)"],
      ["허가 관할", "무석시 응급관리국·행정심사국 → 구 단위 접수 (확인 필요)"],
      ["고객 접근", "SK하이닉스 우시 약 3 km · CXMT 약 296 km (직선)"],
      ["핵심 관문", "원구 내 3PL이 우리 품목(자연발화·수분반응 가능성)을 받아주는가"],
      ["미확인", "신오구 처리기한 · 무저장 장소 요건 · 임차저장이 관할 밖 창고도 되는지"]], "추천"),
    ("CITY-002", "강소성 후보 2 · 쑤저우: SIP 법인 + 장자강 위험품 물류", "장자강에 위험화학품 부두·보세 위험품 창고 보도 · 쑤저우 세칙은 미확보",
     [["권장 운영 구조", "E안 (SIP 법인 + 장자강 3PL) 또는 장자강 단일 거점"],
      ["법인 주소", "쑤저우공업원구(SIP) 또는 장자강보세구"],
      ["창고 후보", "江苏扬子江国际化学工业园 (장자강) · 常熟 新材料产业园 (불소화학)"],
      ["허가 관할", "쑤저우시 / 장자강시 (현급시) 응급관리 부문 — 확인 필요"],
      ["고객 접근", "SK하이닉스 우시 약 40 km · CXMT 약 332 km (직선)"],
      ["핵심 관문", "보세 위험품 창고의 품목·물성 범위, 법인과 창고 관할 분리 가능 여부"],
      ["미확인", "쑤저우 경영허가 세칙 · 장자강 창고 운영사·견적"]], "조건부"),
    ("CITY-003", "강소성 후보 3 · 창저우: 중장기 자체시설 후보", "무석·난징 사이 인정 원구(滨江 新材料产业园) · 공개 운영 자료는 부족",
     [["권장 운영 구조", "초기는 B안(무석 법인) 유지, 중장기 D안 자체 창고 후보지"],
      ["법인 주소", "초기에는 두지 않음 (무석 법인의 창고 거점 후보)"],
      ["창고 후보", "常州滨江经济开发区新材料产业园 (인정 원구)"],
      ["허가 관할", "창저우시 응급관리 부문 · 원구 관리위원회"],
      ["고객 접근", "SK하이닉스 우시 약 67 km · CXMT 약 252 km (직선)"],
      ["핵심 관문", "외자 소규모 무역법인의 원구 입주·창고 임차 조건"],
      ["미확인", "원구 위험품 창고 현황·임차료·입주 기준"]], "확인 필요"),
]
for cid, ttl, sub, rows_, lab in DET:
    s = slide(ttl.split(" · ")[1], sub)
    tb(s, 0.31, 1.05, 6.0, 0.35, [[(ttl.split(" · ")[0], 13, True, REDC)]])
    sc = SCORE[cid]
    badge(s, 7.0, 1.06, 1.25, 0.32, lab)
    tb(s, 8.3, 1.07, 1.4, 0.32, [[(f"점수 {sc[0]:.0f}" if isinstance(sc[0], (int, float)) else "점수 보류", 11, True, DARK)]], align="r")
    table(s, 0.31, 1.5, [1.75, 7.62], [["항목", "내용"]] + rows_, fs=12, rh=[0.32] + [0.6] * 7, left_cols=(0, 1))
    src(s, "08·09 시트 (CITY·SITE), EVD-014·020·021 · 직선거리 대략, 점수 분석 가정")
    notes(s, f"{CN[cid][2]} 후보 상세. 같은 항목·같은 순서로 3개 후보를 비교할 수 있게 했습니다. 판정 '{lab}'의 근거: {CN[cid][13]}. 공개 근거로 확인된 것은 원구의 '인정' 지위까지이며, 창고 운영사·품목·견적은 모두 문의 예정입니다.")

# ================================================================== 18 강소성 추천·조건·탈락
s = slide("강소성 결론: 무석 1순위, 쑤저우 차선 (모두 조건부)", "추천을 바꿀 수 있는 조건 = 3PL 품목 수용 · 관할 회신 · 제품 분류")
rows = [["순위", "후보", "운영 구조", "확정 조건", "탈락·보류 이유"],
        ["1순위", "무석 신오구", "B안 → C안", "관할 원구 3PL이 품목 수용 · 신오구 기한 회신", "-"],
        ["차선", "쑤저우 SIP + 장자강", "E안", "보세 위험품 창고 품목 · 관할 분리 허용", "세칙 미확보"],
        ["3순위", "창저우 滨江", "D안 후보지", "원구 입주 조건", "공개 자료 부족"],
        ["보류", "난징·전장·난퉁·양저우·타이저우", "-", "-", "고객·수입항 동선 불리 / 자료 부족"],
        ["제외", "쿤산 (창고)", "사무만 가능", "-", "1차 명단에 원구 없음"]]
table(s, 0.31, 1.08, [0.9, 2.3, 1.3, 3.0, 1.87], rows, fs=11.5, rh=[0.32] + [0.6] * 5, left_cols=(1, 2, 3, 4), fills={(1, -1): "FFF2E6"})
sec(s, 4.55, "법인 주소와 창고는 같은 도시·같은 성이어야 하나?")
bullets(s, 0.31, 4.88, 9.37, 1.7, [
    "무저장 경영은 사무실 주소 관할에서 허가 — 물건은 허가 창고(3PL)에 있으므로 창고는 다른 도시도 가능성 있음",
    "임차저장 경영은 임차 창고가 허가 내용에 들어가므로 관할(시·성) 제한이 있을 수 있음 → 무석 질의 항목",
    "운송: 성 경계를 넘는 위험물 도로운송 자체는 일반적이나 차량·경로 제한은 운송사 확인"], fs=12)
src(s, "08_CityComparison, 09_SitesWarehouses · EVD-014·015")
notes(s, "강소성 안에서 실행 가능성이 가장 높은 조합은 무석 법인 + 무석 관할 원구 3PL입니다. 법인과 창고를 다른 도시(또는 다른 성)에 두는 E안은 무저장 경영이면 성립 가능성이 높지만, 임차저장 경영으로 가면 관할 문제가 생길 수 있어 기관 질의가 필요합니다.")

# ================================================================== 19 장애 유형
s = slide("강소성에서 막혀도 국가 규정 문제는 이전으로 풀리지 않습니다", "지역을 바꿔 해결되는 것은 '창고·원구 진입'과 '처리 속도' 정도")
rows = [["장애 유형", "성격", "다른 지역으로 해결?", "전환 조건"],
        ["신규물질 등록 명의 (개정안)", "국가 규정", "안 됨", "-"],
        ["위험화학품 경영허가 필요", "국가 규정", "안 됨 (지역별 세칙·기한만 다름)", "-"],
        ["저장은 지정 구역만", "국가·성", "부분 (원구 여유·수용 품목이 다름)", "강소성 원구 3PL 전부 불가 시"],
        ["특정 원구·창고 입주 거절", "단지 기준", "해결 가능", "후보 3곳 이상 거절 시"],
        ["자연발화·수분반응 품목 미수용", "제품·창고", "해결 가능 (전문 창고 찾기)", "물성 확정 후 수용 창고 없음"],
        ["처리기한·비용 과다", "지방 실무", "일부 해결", "무석 회신이 3개월 이상"]]
table(s, 0.31, 1.08, [2.75, 1.4, 2.85, 2.37], rows, fs=12, rh=[0.32] + [0.55] * 6, left_cols=(0, 2, 3), fills={(1, -1): "F2F2F2", (2, -1): "F2F2F2"})
tb(s, 0.31, 5.0, 9.37, 1.5, [[("정리  ", 12, True, REDC), ("상하이·저장·안후이로 옮겨도 등록 명의·경영허가 의무는 똑같습니다. 이전 검토는 '창고를 못 찾을 때'와 '고객(CXMT) 대응'에서만 의미가 있습니다.", 12, False, DARK)],
                               [("유지안  ", 12, True, REDC), ("강소성 법인을 두고 창고만 타 지역 3PL을 쓰는 방안(E안)은 무저장 경영이면 성립 가능성 — 무석 질의로 확인.", 12, False, DARK)]])
src(s, "06_LegalMatrix · EVD-002·008·013")
notes(s, "장애를 국가 규정·성 공통·단지 기준·제품·비용으로 나눴습니다. 국가 규정 문제는 지역 이전으로 회피되지 않으므로, 대안 지역 검토는 창고 확보와 고객 거리 문제에 한정합니다.")

# ================================================================== 20 대안 지도·비교
s = slide("인근 대안은 상하이(수입항)와 허페이(CXMT)입니다", "저장성 자싱은 창고 대안 · 닝보는 고객 거리로 제외")
pts = [("무석 (1순위)·SK하이닉스", 120.37, 31.55, RED, -1.75, -0.2, True), ("상하이 외고교", 121.60, 31.36, BLU, -1.0, 0.08, True),
       ("상하이 화학공업구", 121.45, 30.80, BLU, -1.3, -0.12, False), ("자싱·핑후", 121.10, 30.62, "5B9BD5", -0.85, 0.02, False),
       ("닝보 진해 (제외)", 121.66, 29.98, "A6A6A6", -1.2, -0.12, False), ("허페이 경개구·CXMT", 117.24, 31.75, BLU, 0.12, -0.12, True),
       ("추저우", 118.32, 32.30, "A6A6A6", 0.1, -0.12, False), ("우후·마안산", 118.47, 31.50, "A6A6A6", 0.1, -0.1, False),
       ("장자강", 120.43, 31.95, ORG, 0.1, -0.15, False)]
H20 = draw_map(s, 0.31, 1.08, 5.4, {"江苏省"}, {"浙江省", "安徽省", "上海市"}, pts, [("", 120.42, 31.47, 0, 0), ("", 117.32, 31.68, 0, 0)])
rows = [["대안", "해결하는 것", "판정"], ["상하이 (직할시)", "수입항·보세 물류·자유무역구 처리 속도", "조건부"], ["허페이 (안후이)", "CXMT 대응 (직선 약 1 km)", "확인 필요"],
        ["자싱·핑후 (저장)", "강소성 창고 미확보 시 원구 창고", "보류"], ["닝보 (저장)", "-", "제외"]]
table(s, 5.85, 1.08, [1.35, 1.6, 0.88], rows, fs=10.5, rh=[0.3, 0.62, 0.62, 0.62, 0.45], status_col=2, left_cols=(0, 1))
tb(s, 5.85, 3.85, 3.83, 2.7, [[("국가 규정은 동일, 다른 것은", 11.5, True, REDC)],
                              [("· 원구·창고의 여유와 수용 품목", 10.5, False, DARK)], [("· 처리기한 (자유무역구 무저장 고지승낙 15 영업일 사례)", 10.5, False, DARK)],
                              [("· 고객·항만까지 거리", 10.5, False, DARK)], [("YMTC(우한)는 지도 밖 서쪽 — 모든 후보에서 450 km 이상", 10.5, False, GRAY)]])
src(s, "SRC-031·032 상하이, SRC-033 저장 원구, SRC-034 안후이 원구, EVD-016 · 지도 경계: 공개 GeoJSON(단순화)")
notes(s, "상하이는 성이 아닌 직할시입니다. 법인·수입 관리 거점으로는 강점이 있지만 비용이 높고, 위험품 저장은 시가 정한 구역(화학공업구 등)에서만 가능합니다. 허페이는 CXMT와 가깝지만 허페이 관할 화공원구·위험품 창고 공개 근거를 찾지 못해 '확인 필요'입니다. 자싱(핑후 독산항·자싱항구)은 창고 대안입니다.")

# ================================================================== 21 인근 대안 상세
s = slide("상하이는 수입 거점, 허페이는 CXMT 거점으로 의미가 있습니다", "두 곳 모두 '강소성 법인 유지 + 보완 거점'이 우선 — 단독 이전은 차순위")
rows = [["항목", "인근 대안 1 · 상하이 (직할시)", "인근 대안 2 · 허페이 (안후이성)"],
        ["역할", "수입 통관·보세 물류 거점 (외고교·임항)", "CXMT 대응 영업·재고 거점"],
        ["권장 구조", "E안: 무석 법인 + 상하이 3PL / 또는 상하이 법인", "무석 법인 + 허페이 3PL / 필요 시 지점"],
        ["창고 후보", "상하이 화학공업구·항만 보세 창고 (확인)", "공개 근거 없음 (정보 없음)"],
        ["허가 근거", "상하이 위험화학품 안전관리 규장 · 창고 품목 추가 지침", "안후이 원구 39곳 인정 (허페이 포함 여부 미확인)"],
        ["고객 거리", "SK하이닉스 약 116 km · CXMT 약 414 km", "CXMT 인접 · SK하이닉스 약 300 km"],
        ["약점", "임대·인건비 높음 · 저장 구역 제한", "창고 후보 부재 · 강소성 우선 원칙과 거리"],
        ["판정", "조건부", "확인 필요"]]
table(s, 0.31, 1.08, [1.4, 4.0, 3.97], rows, fs=12, rh=[0.32] + [0.6] * 7, left_cols=(0, 1, 2), status_col=None)
src(s, "SRC-031·032 상하이시, SRC-034 안후이성 응급관리청 · 직선거리 대략(12_LogisticsStaff)")
notes(s, "인근 대안도 국가 규정은 같습니다. 상하이는 수입항과 보세 물류가 강점이고, 허페이는 CXMT 납품 대응이 강점입니다. 둘 다 강소성 법인을 유지한 채 보완 거점으로 쓰는 방안을 먼저 보고, 강소성에서 창고를 전혀 못 구할 때 단독 이전을 검토합니다.")

# ================================================================== 22 전체 평가·민감도
s = slide("가중치를 바꿔도 무석이 1순위로 유지됩니다", "2위는 기준에 따라 쑤저우·상하이가 바뀜 — 점수는 분석 가정")
order_ids = ["CITY-001", "CITY-002", "CITY-101", "CITY-003", "CITY-006"]
labs = {"CITY-001": "무석", "CITY-002": "쑤저우+장자강", "CITY-101": "상하이", "CITY-003": "창저우", "CITY-006": "난징"}
rank = []
for k in range(3):
    vals = sorted([(SCORE[c][k], c) for c in order_ids], reverse=True)
    rank.append({c: i + 1 for i, (_, c) in enumerate(vals)})
rows = [["후보", "기본 (허가 우선)", "속도 우선", "비용 우선", "근거 충족"]]
for c in order_ids:
    rows.append([labs[c]] + [f"{SCORE[c][k]:.0f}  ({rank[k][c]}위)" for k in range(3)] + [CN[c][11]])
rows.append(["그 외 11곳", "점수 보류", "점수 보류", "점수 보류", "관문·정보 부족"])
table(s, 0.31, 1.08, [2.2, 1.9, 1.9, 1.9, 1.47], rows, fs=12, rh=[0.32] + [0.45] * 6, left_cols=(0,), fills={(1, -1): "FFF2E6"})
rows = [["가중치", "허가", "창고", "일정", "고객", "비용", "인력"]] + [[k] + [str(x) for x in v] for k, v in D.WEIGHTS.items()]
table(s, 0.31, 4.25, [2.2, 1.19, 1.19, 1.19, 1.19, 1.19, 1.22], rows, fs=11, rh=[0.3] + [0.34] * 3, left_cols=(0,))
rect(s, 0.31, 5.72, 9.37, 0.85, "FFF2E6")
tb(s, 0.45, 5.76, 9.1, 0.8, [[("최종 권고 (조건부)  ", 12, True, REDC), ("강소성 1순위 무석(B안) · 강소성 차선 쑤저우+장자강(E안) · 인근 대안 1 상하이(수입 거점) · 인근 대안 2 허페이(CXMT 거점)", 12, False, DARK)]])
src(s, "08_CityComparison (SUMPRODUCT 수식, 가중치 입력 셀) · 가중치는 제안값이며 공식 기준 아님")
notes(s, "필수 관문을 통과한 5곳만 점수를 매겼고, 나머지 11곳은 정보 부족으로 점수를 보류했습니다(임의 평균 점수 없음). 허가 우선·속도 우선·비용 우선으로 가중치를 바꿔도 무석이 1위이며, 2위는 기준에 따라 쑤저우 또는 상하이로 바뀝니다. 점수 자체가 분석 가정이므로 기관 회신 후 갱신해야 합니다.")

# ================================================================== 23 MP vs TC
s = slide("MP가 허가를 이미 가졌는지가 비교의 분기점입니다", "MP 정보(법인명·허가·창고)가 없어 두 시나리오로 비교")
rows = [["항목", "시나리오 1: MP가 허가·창고 보유", "시나리오 2: MP도 새로 취득", "TC 독자 법인"],
        ["속도", "MP 가장 빠름", "TC와 비슷", "등기+허가 필요"],
        ["추가 비용", "MP 수수료", "MP에도 같은 투자", "TC 신규 투자"],
        ["등록 명의", "MP", "MP", "TC"],
        ["자료·IP 통제", "MP 경유", "MP 경유", "TC 직접"],
        ["우선순위·의사결정", "MP 내부에 종속", "MP 내부에 종속", "TC 직접"],
        ["계약 종료 시", "등록증·고객 이관 문제", "동일", "해당 없음"],
        ["합리적 결론", "단기 MP 활용 + TC 명의 등록 병행", "TC 독자 법인", "-"]]
table(s, 0.31, 1.08, [1.75, 2.75, 2.45, 2.42], rows, fs=11.5, rh=[0.32] + [0.47] * 7, left_cols=(0,), fills={(7, -1): "FFF2E6"}, bolds=(7,))
sec(s, 4.95, "독자 법인 명분과 반론")
bullets(s, 0.31, 5.25, 9.37, 1.35, [
    "명분: 개정안 확정 시 등록 명의가 곧 사업권 — 그룹 법인이라도 법인이 다르면 명의·책임·이전가격이 갈림",
    "반론: MP가 해당 허가를 이미 가졌다면 시간·비용은 MP가 유리 → 'MP는 허가 미보유' 회의 언급부터 확인"], fs=12)
src(s, "11_MPComparison 시트 · EVD-002·005 · MP 관련 항목은 정보 없음/내부 제공")
notes(s, "MP의 정식 법인명·사업범위·허가·창고·등록 실적을 모르는 상태라 독자 법인 우위를 단정하지 않았습니다. 시나리오 2(MP도 새로 취득)라면 같은 투자를 MP에 하는 셈이므로 TC 독자 법인이 명의·통제 측면에서 합리적입니다. 시나리오 1이라면 단기에는 MP를 쓰되 등록은 TC 명의로 진행하는 혼합안이 현실적입니다. '그룹 법인은 비밀 유지가 된다', 'TC 법인이 무조건 빠르다'는 가정은 쓰지 않았습니다.")

# ================================================================== 24 물류·조직
s = slide("물류는 '수입항 → 허가 창고 → 고객' 한 줄로 설계합니다", "재고일수는 미제공 — 1·2·3개월 민감도로만 비교")
chev(s, 0.31, 1.12, 9.37, 0.82, [("한국 TC", "제조·출하", "D9D9D9", DARK), ("수입항", "상하이·장자강", BLU, "FFFFFF"), ("허가 창고", "3PL (원구)", GRN, "FFFFFF"),
                                  ("고객 공장", "무석·허페이", ORG, "FFFFFF"), ("공병 회수", "반송·세정", "7F7F7F", "FFFFFF")], fs=11.5)
rows = [["확인 항목", "내용"],
        ["항만·운송", "해당 위험물 반입 가능 항만, 위험물 운송 자격 차량·경로"],
        ["3PL 조건", "품목 허가·불활성가스·온습도·보험·출입 통제"],
        ["재고 (민감도)", "1·2·3개월분 = 월 판매량 × 1·2·3 (판매량 미제공)"],
        ["용기 회수", "잔류물·세정·국제 반송은 별도 규제 — 충진·세정은 F안"]]
table(s, 0.31, 2.12, [1.6, 7.77], rows, fs=12, rh=[0.32] + [0.45] * 4, left_cols=(0, 1))
sec(s, 4.45, "운영 조직 (최소 vs 확대)", "12_LogisticsStaff")
rows = [["역할", "최소", "확대", "비고"], ["법인장·법정대표인", "1 (주재원)", "1", "경영허가 '주요 책임자' 자격"], ["전담 안전관리자", "1", "1", "겸직 불가 해석 주의"],
        ["영업·기술 / 경리", "1 / 1", "2 / 1", "경리 외주 가능"], ["물류·통관 / 창고·품질", "0 / 0", "1 / 2", "C·D안에서 필요"]]
table(s, 0.31, 4.75, [2.6, 1.2, 1.2, 4.37], rows, fs=11.5, rh=[0.3] + [0.38] * 4, left_cols=(0, 3))
src(s, "EVD-012·019·023 · 12_LogisticsStaff · 직선거리·운송시간은 운송사 확인 전")
notes(s, "도로거리·실제 운송 리드타임은 위험물 차량 통행 제한·출입 절차가 반영돼야 하므로 지도 시간으로 대체하지 않았습니다. 현지 재고는 기존 거래처(유통사)의 역할·수수료에도 영향을 주므로 영업 측 검토가 필요합니다.")

# ================================================================== 25 비용
c_init, c_dep, c_fix, c_var, c_wc = (COST[k] for k in ("초기", "보증금", "연간", "변동", "운전자본"))
s = slide("확인 가능한 비용은 인건비뿐이고 나머지는 견적이 필요합니다", f"B안 현지 인건비 연 {PERS[0][3]:,.0f} CNY ≈ {PERS[0][4]:,.0f}백만원 (평균임금 기반 가정)")
rows = [["구분 (B안)", "확인 가능 금액 (CNY)", "KRW (백만원)", "상태", "주요 항목"],
        ["초기 일회성", f0(c_init[0]), f0(c_init[3]), c_init[2], "설립 대행·법무·등록 시험·교육·IT"],
        ["보증금 (회수 가능)", f0(c_dep[0]), f0(c_dep[3]), c_dep[2], "사무실 보증금 3개월 (가정)"],
        ["연간 고정비", f0(c_fix[0]), f"{c_fix[3]:,.0f}", c_fix[2], "인건비(산출)·임대·회계·보험·허가 유지"],
        ["변동비 (연)", f0(c_var[0]), f0(c_var[3]), c_var[2], "3PL 보관·입출고·위험물 운송"],
        ["운전자본", f0(c_wc[0]), f0(c_wc[3]), c_wc[2], "재고 = 월 판매액 × 재고일/30"]]
table(s, 0.31, 1.08, [1.8, 1.75, 1.2, 1.65, 2.97], rows, fs=11.5, rh=[0.32] + [0.48] * 5, left_cols=(0, 4))
rows = [["인건비 시나리오", "현지 인원", "연 인건비 (CNY)", "KRW (백만원)"]] + [[p[0], f"{p[1]:.0f}명", f0(p[3]), f"{p[4]:,.0f}"] for p in PERS]
table(s, 0.31, 4.12, [3.0, 1.5, 2.5, 2.37], rows, fs=11.5, rh=[0.3] + [0.36] * 3, left_cols=(0,))
tb(s, 0.31, 5.6, 9.37, 1.0, [[("구분  ", 11.5, True, REDC), (f"등록자본(법정 최저 없음, 5년 내 납입) ≠ 투자비 ≠ 보증금 ≠ 재고자금. 등록자본 제안 = 확인 가능 고정비 × 1.5 = {CAPV[0]:,.0f} CNY (미완성 — 견적 반영 후 재산정). "
                                                          f"환율 1 CNY = {FX:.0f} KRW (2026-09-01, 갱신 필요). 주재원(법인장) 비용은 본사·법인 부담 결정 전.", 11, False, DARK)]])
src(s, "15_CostModel·14_CostsRaw·13_FXAssumptions · EVD-027 무석 평균임금 147,293 CNY, EVD-034 환율, EVD-036 사회보험 기수")
notes(s, f"""확인 가능 금액만 합산했고 견적이 필요한 항목은 빈칸으로 두어 '미완성(견적 n건)'으로 표시했습니다. 미확인 비용을 0원으로 처리하지 않았습니다.
인건비 = 무석 2025 평균임금 147,293 CNY × 직무 배수 1.2 × (1 + 사업주 부담 35%) — 배수·부담률은 분석 가정입니다.
자금 회수 경로: 배당(법정 공적금 적립 후, 원천징수 10% — 한중 조약 세율 확인), 본사 서비스 대가(이전가격), 차입금 상환, 감자·청산. '중국 자금은 돌려받을 수 없다'는 단정은 근거가 없습니다(EVD-035).
최소 비용안(B) vs 중장기 적정안(C): 인건비 시나리오 '확대'가 C안에 해당하며 창고 임차·안전 설비가 추가됩니다.""")

# ================================================================== 26 일정
s = slide("연말까지 가능한 것은 법인 등기와 기존 备案 전환 신청입니다", "기준 시나리오: 허가 2월 · TC 명의 등록 5월 · 초도 공급 2027년 7월")
GX, GY, GW = 2.75, 1.5, 6.9
d0, d1 = dt.date(2026, 10, 1), dt.date(2028, 1, 1)
span = (d1 - d0).days
def gx(dd):
    dd = dd.date() if isinstance(dd, dt.datetime) else dd
    return GX + (dd - d0).days / span * GW
m = d0
while m < d1:
    x = gx(m)
    rect(s, x, GY - 0.32, 0.005, 4.15, None, "D9D9D9")
    tb(s, x, GY - 0.36, 0.5, 0.25, [[(f"{m.month}" if m.month != 1 else f"'{m.year % 100}.1", 8, True, GRAY)]])
    m = dt.date(m.year + (m.month // 12), m.month % 12 + 1, 1)
gant = ["T01", "T03", "T05", "T07", "T09", "T10", "T11", "T12", "T13", "T14", "T20"]
LBL = {"T01": "물질·규제 분류", "T03": "기관 질의", "T05": "경영진 승인", "T07": "법인 등기", "T09": "무저장 경영허가", "T10": "위험화학품 등록",
       "T11": "신규물질 등록 (TC 명의)", "T12": "3PL 계약", "T13": "고객 공급자 승인", "T14": "초도 공급", "T20": "연말: 备案 전환 신청"}
for i, tid in enumerate(gant):
    t = TASK[tid]; y = GY + i * 0.36
    tb(s, 0.31, y - 0.02, 2.4, 0.3, [[(LBL[tid], 10.5, tid in ("T07", "T11", "T14", "T20"), REDC if tid == "T20" else DARK)]])
    x0, x1 = gx(t["so"]), gx(t["ep"])
    rect(s, x0, y + 0.06, max(x1 - x0, 0.03), 0.18, "F2F2F2", "BFBFBF")
    xs, xe = gx(t["s"]), gx(t["e"])
    rect(s, xs, y + 0.06, max(xe - xs, 0.04), 0.18, REDC if tid == "T20" else (ORG if tid in ("T11", "T09") else BLU))
yd = gx(dt.date(2026, 12, 31))
rect(s, yd, GY - 0.32, 0.02, 4.15, REDC)
tb(s, yd - 0.45, GY + 3.86, 0.9, 0.22, [[("12-31", 9, True, REDC)]], align="c")
tb(s, 0.31, 5.72, 9.37, 0.95, [[("읽는 법  ", 11, True, REDC), ("진한 막대 = 기준 시나리오, 옅은 막대 = 낙관~보수 범위 (기간은 분석 가정, 공휴일 미반영). ", 10.5, False, DARK)],
                               [("기준 시나리오  ", 11, True, REDC), (f"법인 등기 {md(TASK['T07']['e'])} · 경영허가 {md(TASK['T09']['e'])} · TC 명의 등록 {md(TASK['T11']['e'])} · 초도 공급 {md(TASK['T14']['e'])}  |  "
                                                                   f"备案 전환 신청 {md(TASK['T20']['e'])} (낙관 {md(TASK['T20']['eo'])} · 보수 {md(TASK['T20']['ep'])})", 10.5, False, DARK)]])
src(s, "16_Timeline (선행관계 수식, 3시나리오) · EVD-003 연말 기한 · EVD-016 처리기한 사례")
notes(s, f"""착수 2026-10-06 기준입니다. 법인 등기 완료(기준 {d8(TASK['T07']['e'])})는 사업 개시가 아닙니다. 수입·판매는 경영허가({d8(TASK['T09']['e'])}), 위험화학품 등록({d8(TASK['T10']['e'])}), 신규물질 등록({d8(TASK['T11']['e'])}) 이후입니다.
임계경로: 법인 등기 → TC 명의 신규물질 등록(중국 내 시험 포함 가정 20주) → 고객 승인 → 초도 공급.
연말 목표: TC 법인 명의로 12-31까지 등록을 끝내는 것은 어렵습니다. 가능한 합법적 대응은 ① 기존 备案 보유자가 12-31 전 등록증 전환 신청(기준 {d8(TASK['T20']['e'])}) ② 현행 규정이 유지되면 TC 본사+대리인 신청 ③ 적격 수입자 경유 공급 ④ 고객에 일정 사전 안내. 기관 재량 유예는 근거·사례가 없어 일정에 넣지 않았습니다.
기간은 분석 가정이며 국경절·춘절 등 공휴일은 반영하지 않았습니다.""")

# ================================================================== 27 리스크·2주 과제
s = slide("첫 2주에 결론을 바꿀 사실 5가지를 확인합니다", "备案 보유자 · 제품 분류 · 개정안 확정 · 3PL 품목 · MP 허가")
rows = [["리스크", "영향", "조기 확인", "담당", "기한"]]
for rid in ("R02", "R01", "R03", "R05", "R13", "R04", "R10"):
    r_ = next(x for x in D.RISKS if x[0] == rid)
    rows.append([r_[1][:30], r_[2], r_[7][:26], r_[5], r_[8][5:]])
table(s, 0.31, 1.08, [3.3, 0.75, 3.1, 1.12, 1.1], rows, fs=10.5, rh=[0.3] + [0.47] * 7, left_cols=(0, 2))
sec(s, 4.75, "2주 실행 계획 (10/6~10/20)")
bullets(s, 0.31, 5.05, 9.37, 1.6, [
    "1주: 제품별 CAS·SDS·수입량·备案 번호 수집(한호진 PL) · MP 실사 질문지 발송(기획) · 법무 전문가 후보 선정",
    "1~2주: 무석 신오구·쑤저우 응급관리/행정심사 질의, 생태환경부 등록 상담, 3PL 3곳(무석 원구·장자강·상하이) 품목 문의",
    "2주 말: 판정표·회신 정리 → 구조·지역 1차 확정안 보고 (30/60/90일 계획은 부록)"], fs=11.5)
src(s, "17_RisksActions 시트 (R01~R13) · 모든 기관 문의 상태: 문의 예정")
notes(s, "리스크 13개 중 결론을 바꿀 수 있는 것을 먼저 확인합니다. 부서별 역할: 해외영업(명분·제품·고객), 김지현 PL(규정 1차 조사), 한호진 PL(제품·공급 구조), 법무(법령·계약·현지 전문가), 여현석 팀장 등 기획(투자·운전자본), 인사(인력·채용), 중국 현지(입지·시설·업체). 기관에 실제 연락한 것은 없으며 모두 '문의 예정'입니다.")

# ================================================================== 28 의사결정 요청
s = slide("오늘 결정할 것은 '검증 착수'이고 확정은 11월입니다", "현재 최종 도시·구조 확정 불가 — 병행 추진 후보를 승인 요청")
rows = [["결정 항목", "제안", "조건", "기한"],
        ["운영 구조", "B안 (TC 법인 + 무저장 + 3PL) 우선, MP·대리인안 병행 검토", "MP 허가 보유 시 재검토", "11월 중순"],
        ["지역", "무석 1순위 · 쑤저우 차선 · 상하이/허페이 보완", "3PL 품목 수용 회신", "11월 중순"],
        ["연말 대응", "备案 보유자 확인 → 12-31 전 전환 신청 지원", "TC 제품이 备案 대상일 때", "10/16 확인"],
        ["예산", "2주 확인 비용(전문가 자문) 승인 · 본 예산은 견적 후", "견적 수령", "10/20"],
        ["조직", "전담 안전관리자 채용 착수 · 법인장 지정 검토", "구조 확정 시", "11월"],
        ["다음 단계", "2주 확인 결과 보고 → 구조·지역·예산 확정", "-", "10/20"]]
table(s, 0.31, 1.08, [1.4, 4.4, 2.35, 1.22], rows, fs=12, rh=[0.32] + [0.62] * 6, left_cols=(0, 1, 2), fills={(1, -1): "FFF2E6"})
tb(s, 0.31, 5.35, 9.37, 1.2, [[("확정에 필요한 조건  ", 12, True, REDC), ("① 개정 12호령 최종 내용 ② TC 제품의 신규·위험 여부와 备案 보유자 ③ 무석(또는 쑤저우) 허가 관할·기한 회신 "
                                                                  "④ 3PL 1곳 이상 품목 수용 확인 ⑤ MP 허가·창고 보유 여부", 12, False, DARK)]])
src(s, "02_ExecutiveSummary · 17_RisksActions")
notes(s, "추천을 억지로 확정하지 않았습니다. 오늘은 검증 착수와 병행 후보를 승인받고, 2주 확인 결과로 11월 중순 구조·지역·예산을 확정하는 일정입니다.")

# ================================================================== 부록
def appendix(title, sub):
    s_ = slide(title, sub)
    return s_


s = appendix("부록 1 | 핵심 용어를 쉬운 말로 풀었습니다", "같은 말처럼 들려도 등록·허가·명의는 서로 다른 절차입니다")
rows = [["용어", "쉬운 뜻", "이번 사업과의 관계"]] + [[g[0], g[1], g[2]] for g in D.GLOSSARY[:9]]
table(s, 0.31, 1.08, [2.5, 4.0, 2.87], rows, fs=10.5, rh=[0.3] + [0.56] * 9, left_cols=(0, 1, 2))
src(s, "19_Glossary_Q 시트")
notes(s, "용어 해설 (1/2).")
s = appendix("부록 1 | 핵심 용어 (계속)", "법인장과 법정대표인은 같은 사람일 수도, 다를 수도 있습니다")
rows = [["용어", "쉬운 뜻", "이번 사업과의 관계"]] + [[g[0], g[1], g[2]] for g in D.GLOSSARY[9:]]
table(s, 0.31, 1.08, [2.5, 4.0, 2.87], rows, fs=11, rh=[0.3] + [0.62] * 5, left_cols=(0, 1, 2))
src(s, "19_Glossary_Q 시트")
notes(s, "용어 해설 (2/2).")

s = appendix("부록 2 | 법령·허가 근거표", "결론은 근거 기반 판단이며 법률 의견이 아닙니다 — 원문 대조 필요")
rows = [["쟁점", "규정", "결론", "근거 ID"]] + [[l[0], l[2], l[8], l[11]] for l in D.LEGAL]
table(s, 0.31, 1.08, [2.6, 2.1, 3.6, 1.07], rows, fs=9.5, rh=[0.28] + [0.5] * 10, left_cols=(0, 1, 2))
src(s, "06_LegalMatrix 시트 · 03_SourceRegister")
notes(s, "06_LegalMatrix 요약. 상세 조건·미확인 사항은 Excel에 있습니다.")

s = appendix("부록 3 | 후보 단지·창고 전체 조사표", "공개 근거는 '인정 원구' 지위까지 — 품목·물성·견적은 모두 문의 예정")
rows = [["ID", "단지·창고", "공식 인정", "품목·특수 물성", "상태"]] + [[x[0], x[2][:34], x[6], f"{x[7][:14]} / {x[8][:10]}", x[12]] for x in D.SITES]
table(s, 0.31, 1.08, [0.85, 3.9, 1.6, 2.1, 0.92], rows, fs=9.5, rh=[0.28] + [0.38] * 13, left_cols=(1, 2, 3))
src(s, "09_SitesWarehouses 시트 · EVD-020·021 · SRC-031~034")
notes(s, "SITE-001~013. 공개 연락처·주소는 기관 공식 사이트에서 확인 후 기록 예정입니다. 허위 담당자·견적을 만들지 않았습니다.")

s = appendix("부록 4 | 제품별 입력자료와 규제 판단", "제품 정보가 들어오면 이 표가 시나리오 A~D 중 하나로 확정됩니다")
rows = [["제품", "신규물질", "위험화학품", "운송", "물성·보관", "기존 备案", "필요 자료"]] + [[m[1][:14], "확인 필요", "확인 필요", "UN 확인", "확인 필요", "확인 필요", "SDS·CAS·물량·备案 번호"] for m in D.MATERIALS]
table(s, 0.31, 1.08, [2.0, 1.05, 1.15, 0.95, 1.15, 1.1, 1.97], rows, fs=10.5, rh=[0.3] + [0.5] * 4, left_cols=(0, 6))
sec(s, 3.6, "판정 순서")
chev(s, 0.31, 3.95, 9.37, 0.8, [("CAS 확인", "IECSC 검색", BLU, "FFFFFF"), ("GHS 분류", "SDS", BLU, "FFFFFF"), ("목록 대조", "위험화학품", ORG, "FFFFFF"),
                                  ("UN 분류", "운송", ORG, "FFFFFF"), ("물량·备案", "등록 유형", GRN, "FFFFFF")], fs=11)
src(s, "05_Materials 시트 · EVD-032")
notes(s, "CAS·위험등급을 임의로 넣지 않았습니다.")

s = appendix("부록 5 | 비용 산식·가정·환율", "파란 입력값을 견적으로 바꾸면 15_CostModel이 자동 갱신됩니다")
rows = [["가정 ID", "항목", "값", "상태"]] + [[a[0], a[1], ("견적 필요" if a[2] is None else f"{a[2]:,}" if isinstance(a[2], (int, float)) and a[2] >= 10 else str(a[2])), a[10]] for a in D.ASSUMPTIONS]
table(s, 0.31, 1.08, [1.0, 4.6, 1.6, 2.17], rows, fs=10, rh=[0.28] + [0.37] * 11, left_cols=(1,))
src(s, "13_FXAssumptions · 14_CostsRaw · SRC-024~026")
notes(s, "인건비 = BASE_WAGE × WAGE_MULT × (1+ER_RATE) × LOCAL_HC. 임대료 = OFFICE_RENT × OFFICE_SQM × 12 (단가 견적 필요). 운전자본 = MONTHLY_SALES × INV_DAYS/30 (판매액 미제공). KRW = CNY × FX_KRW.")

s = appendix("부록 6 | 일정 세부표와 시나리오", "달력일·주 단위 · 공휴일 미반영 — 영업일 일정 아님")
rows = [["ID", "단계", "낙관 종료", "기준 종료", "보수 종료"]] + [[tid, LBL.get(tid, TASK[tid]["name"][:22]), md(TASK[tid]["eo"]), md(TASK[tid]["e"]), md(TASK[tid]["ep"])] for tid in TASK]
table(s, 0.31, 1.08, [0.7, 4.6, 1.35, 1.35, 1.37], rows, fs=9.5, rh=[0.28] + [0.335] * len(TASK), left_cols=(1,))
src(s, "16_Timeline 시트 (선행관계 수식)")
notes(s, "선행 작업의 종료 다음 날 시작. 기간(주)은 분석 가정이며 Excel에서 바꿀 수 있습니다.")

for grp_keys, title in ((["A. 중국 법률·등록 전문가", "B. 도시·산업단지 (무석 신오구 등)"], "부록 7 | 기관 질문지 A·B"),
                        (["C. 창고·운송사", "D. MP", "E. SKTC 내부"], "부록 7 | 기관 질문지 C·D·E")):
    s = appendix(title, "모든 질의의 진행 상태: 문의 예정 (실제 연락·회신 없음)")
    y = 1.1
    for g in grp_keys:
        tb(s, 0.31, y, 9.37, 0.28, [[(g, 11.5, True, REDC)]]); y += 0.3
        qs = D.QUESTIONS[g]
        h = 0.36 * len(qs) + 0.1
        bullets(s, 0.31, y, 9.37, h, qs, fs=10, gap=2); y += h + 0.05
    src(s, "19_Glossary_Q 시트")
    notes(s, "질문지는 Excel 19 시트에 원본이 있습니다. 회신이 오면 09·17 시트 상태를 갱신합니다.")

s = appendix("부록 8 | 30·60·90일 실행 계획", "법인 등기 완료는 출발점 — 허가·등록·고객 승인까지 관리")
rows = [["기간", "목표", "산출물", "책임"],
        ["~30일 (11/5)", "구조·지역 확정, 备案 전환 대응 착수", "판정표·기관 회신·법률 의견·경영진 승인", "해외영업·법무·기획"],
        ["~60일 (12/5)", "해외직접투자 신고·법인 등기 신청, 안전관리자 채용", "신고 수리·등기 서류·채용", "기획·법무·인사"],
        ["~90일 (1/4)", "등기 완료·경영허가 신청·3PL 계약·등록 시험 착수", "营业执照·허가 신청·계약·시험 계약", "현지·물류·한호진 PL"]]
table(s, 0.31, 1.08, [1.5, 3.3, 3.2, 1.37], rows, fs=11.5, rh=[0.3, 0.75, 0.75, 0.75], left_cols=(1, 2, 3))
src(s, "16_Timeline · 17_RisksActions")
notes(s, "30/60/90일 계획은 기준 시나리오 일정과 맞췄습니다.")

srcs = D.SOURCES
for part, chunk in ((1, srcs[:20]), (2, srcs[20:])):
    s = appendix(f"부록 9 | 출처 목록 ({part}/2)", "URL·조문·주의사항 전체는 03_SourceRegister 시트")
    rows = [["ID", "문서명", "기관", "시행일", "상태"]] + [[x[0], x[1][:44], x[2][:13], x[7], "공식·요약" if x[4] == "공식" else "비공식"] for x in chunk]
    table(s, 0.31, 1.08, [0.8, 5.1, 1.8, 0.95, 0.72], rows, fs=8.5, rh=[0.26] + [0.262] * len(chunk), left_cols=(1, 2))
    src(s, "03_SourceRegister 시트 (하이퍼링크)")
    notes(s, "출처 목록. 공식 출처라도 원문 조문은 작업 환경 제약으로 직접 열지 못했고 검색 요약으로 확인했습니다.")

s = appendix("부록 10 | 근거 충족 현황", "핵심 주장마다 근거 수준과 남은 확인 사항을 표시합니다")
cnt = {}
for e in D.EVIDENCE:
    cnt[e[7]] = cnt.get(e[7], 0) + 1
rows = [["정보 상태", "근거 건수", "의미"]] + [[k, str(cnt.get(k, 0)), v[:46]] for k, v in D.STATUS_LEGEND]
table(s, 0.31, 1.08, [2.3, 1.0, 6.07], rows, fs=10.5, rh=[0.3] + [0.42] * 7, left_cols=(0, 2))
rows = [["핵심 주장", "근거 수준", "남은 확인"], ["연말 기한(备案 전환)", "공식 통지(요약)", "원문·TC 제품 적용 여부"], ["등록 신청 주체 변경", "공식 초안(요약)", "최종 공포"],
        ["수입·판매에 생산설비 불요", "법률(요약)", "충진·소분 해석"], ["무석 무저장 경영 가능", "지방 세칙(요약)", "신오구 요건·기한"], ["강소성 원구 창고", "성 공고(요약)", "3PL 품목·견적"]]
table(s, 0.31, 4.35, [3.2, 2.6, 3.57], rows, fs=10.5, rh=[0.3] + [0.36] * 5, left_cols=(0, 1, 2))
src(s, "04_EvidenceRaw · 18_PPTTraceability")
notes(s, "근거 36건의 상태 분포와 핵심 주장별 근거 수준입니다.")

prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides (원본 5장 포함 — 정리 전)")

# 원본 템플릿 슬라이드 제거
prs = Presentation(OUT)
lst = prs.slides._sldIdLst
for x in list(lst)[:len(SRC_SLIDES)]:
    prs.part.drop_rel(x.rId); lst.remove(x)
prs.save(OUT)
print("final slides", len(prs.slides))
