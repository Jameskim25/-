# -*- coding: utf-8 -*-
"""CpZr 생산능력 보고서 PPT (v3) — SK trichem 양식(base.pptx)의 공정도·머리글을 재사용하고 수치는 Excel에서 읽는다.

사용: python build_pptx.py <base.pptx> <calc.xlsx(재계산본)> <map.json> <out.pptx>
"""
import copy
import json
import sys

import openpyxl
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from pptx.enum.shapes import MSO_SHAPE
from pptx_helpers import (E, FONT, RGBColor, box, clone, copy_fill, delete, etree, fill, notes, pos, runfmt, set_lines, shapes, tbox,
                          widen_flow)

BASE, XLSX, MAPF, OUT = sys.argv[1:5]
MAP = json.load(open(MAPF))
wb = openpyxl.load_workbook(XLSX, data_only=True)
REF = MAP["REF"]


def V(k):
    s, c = REF[k]
    return wb[s][c].value


VS = {"A": "05_월별_Capa기준", "B": "06_월별_생산계획"}


def lrow(yr, key):
    ws = wb["09_후공정부하"]; r = MAP["LOAD"][str(yr)][key]
    return [ws.cell(row=r, column=2 + m).value for m in range(1, 13)], ws.cell(row=r, column=15).value


def f0(x):
    return f"{x:,.0f}"


def f1(x):
    return f"{x:,.1f}"


def sg(x, d=0):
    return f"{x:+,.{d}f}".replace("-", "−")


DARK, GRAYT, ORG, BLU, GRN, PUR, REDC = "1A1A1A", "7F7F7F", "FF7900", "2E75B6", "548235", "7030A0", "C00000"
prs = Presentation(BASE)
s1, s2, s3, s4, s5 = prs.slides
CHROME = {"TextBox 2", "직사각형 3", "TextBox 5", "타원 6", "그림 7", "직사각형 160", "직선 연결선[R] 162", "TextBox 10"}
SEC = [copy.deepcopy(shapes(s1)["Rectangle 257"]._element), copy.deepcopy(shapes(s1)["TextBox 258"]._element)]


# ------------------------------------------------------------ 슬라이드 · 공통 요소
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
    return new


def keep_chrome(slide, y_from=0.0):
    for sh in list(slide.shapes):
        if sh.name not in CHROME and sh.top / E >= y_from:
            delete(sh)


def new_slide():
    s = dup_slide(s4)
    keep_chrome(s)
    return s


def header(slide, num, ttl, sub):
    d = shapes(slide)
    set_lines(d["TextBox 2"], [ttl]); set_lines(d["TextBox 5"], [sub])
    set_lines(d["타원 6"], [str(num)]); set_lines(d["TextBox 10"], [f"- {num} -"])
    if num >= 10:
        tf = d["타원 6"].text_frame; tf.margin_left = tf.margin_right = 0
        for r_ in tf.paragraphs[0].runs:
            r_.font.size = Pt(7.5)


def sec(slide, y, main, sub="", x=0.31, w=None):
    bar, txt = (copy.deepcopy(e) for e in SEC)
    for el in (bar, txt):
        slide.shapes._spTree.append(el)
        el.xpath("./*[1]/p:cNvPr")[0].set("id", str(max(s_.shape_id for s_ in slide.shapes) + 1))
    b, t = slide.shapes[-2], slide.shapes[-1]
    pos(b, x=x, y=y + 0.03); pos(t, x=x + 0.12, y=y, w=w or (9.2 - x))
    set_lines(t, [[main, f"   {sub}" if sub else ""]])
    runs = t.text_frame.paragraphs[0].runs
    if len(runs) > 1:
        runs[1].font.size = Pt(8.5); runs[1].font.color.rgb = RGBColor.from_string("C00000"); runs[1].font.bold = True
    return t


# ------------------------------------------------------------ 표
def _ln(tcPr, tag, spec):
    ln = etree.SubElement(tcPr, qn(tag))
    if spec:
        ln.set("w", str(spec[1])); sf = etree.SubElement(ln, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr")).set("val", spec[0])
    else:
        ln.set("w", "0"); etree.SubElement(ln, qn("a:noFill"))


def mk_table(slide, x, y, colw, rows, rh=0.2, fs=6.5, merges=(), fills=None, colors=None, bolds=None, wrap=False, left_cols=(0,),
             hdr_rows=1, rh_list=None):
    fills, colors, bolds = fills or {}, colors or {}, bolds or set()
    nr, nc = len(rows), len(colw)
    gf = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(colw)), Inches(rh * nr))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0"); tblPr.set("bandRow", "0")
    for el in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(el)
    for j, w in enumerate(colw):
        tbl.columns[j].width = Emu(int(w * E))
    hts = rh_list or [rh] * nr
    for i in range(nr):
        tbl.rows[i].height = Emu(int(hts[i] * E))
    for r_, a, b in merges:
        tbl.cell(r_, a).merge(tbl.cell(r_, b))
    for i, row in enumerate(rows):
        hdr = i < hdr_rows
        for j in range(nc):
            v = row[j] if j < len(row) else ""
            c = tbl.cell(i, j)
            tf = c.text_frame; tf.word_wrap = wrap
            p = tf.paragraphs[0]
            for r0 in list(p.runs):
                r0._r.getparent().remove(r0._r)
            run = p.add_run(); run.text = "" if v is None else str(v)
            if run.text == "":
                run.text = " "
            col = colors.get((i, j), "FFFFFF" if hdr else DARK)
            if col == DARK and run.text.startswith("−"):
                col = REDC
            runfmt(run, fs, hdr or (i, j) in bolds or (i, -1) in bolds, col)
            p.alignment = PP_ALIGN.LEFT if (j in left_cols and not (hdr and j > 0)) else PP_ALIGN.CENTER
            end = p._p.find(qn("a:endParaRPr"))
            if end is None:
                end = etree.SubElement(p._p, qn("a:endParaRPr"))
            end.set("sz", str(int(fs * 100)))
            c.margin_left = c.margin_right = Emu(36000); c.margin_top = c.margin_bottom = Emu(9000)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tcPr = c._tc.get_or_add_tcPr()
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB", "a:solidFill", "a:noFill"):
                for el in tcPr.findall(qn(tag)):
                    tcPr.remove(el)
            _ln(tcPr, "a:lnL", None); _ln(tcPr, "a:lnR", None)
            _ln(tcPr, "a:lnT", ("404040", 12700) if i == 0 else None)
            _ln(tcPr, "a:lnB", ("404040", 9525) if (hdr and i == hdr_rows - 1) or i == nr - 1 else ("D9D9D9", 6350))
            bg = fills.get((i, j), fills.get((i, -1), "595959" if hdr else "FFFFFF"))
            sf = etree.SubElement(tcPr, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr")).set("val", bg)
    gf.height = Emu(int(sum(hts) * E))
    return tbl


# ------------------------------------------------------------ 30일 운전 Gantt
def hfmt(v):
    return f"{v:.1f}".rstrip("0").rstrip(".")


def cell_paras(cell, paras, anchor=MSO_ANCHOR.TOP):
    """paras: [(runs=[(text, size, bold, color)], align)]"""
    tf = cell.text_frame; tf.word_wrap = True
    ps = tf.paragraphs
    for extra in ps[1:]:
        extra._p.getparent().remove(extra._p)
    for i, (runs, align) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        for r0 in list(p.runs):
            r0._r.getparent().remove(r0._r)
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        for t, size, bold, colr in runs:
            r = p.add_run(); r.text = t; runfmt(r, size, bold, colr)
        end = p._p.find(qn("a:endParaRPr"))
        if end is None:
            end = etree.SubElement(p._p, qn("a:endParaRPr"))
        end.set("sz", str(int(runs[0][1] * 100)))
    cell.vertical_anchor = anchor


def chevrons(slide, x, y, w, h, steps):
    """steps: (단계명, 시간 h, 누적 종료 h, 채움색, 글자색, 대기 여부) — 단계 길이와 무관하게 같은 폭으로 표시."""
    n = len(steps); cw = w / n
    for i, (name, hh, end, colr, tc, wait) in enumerate(steps):
        shp = slide.shapes.add_shape(MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON,
                                     Inches(x + i * cw), Inches(y), Inches(cw + 0.05), Inches(h))
        shp.adjustments[0] = 0.22
        shp.shadow.inherit = False
        fill(shp, colr)
        if wait:
            shp.line.color.rgb = RGBColor.from_string("C00000"); shp.line.width = Pt(1)
            from pptx.enum.dml import MSO_LINE
            shp.line.dash_style = MSO_LINE.DASH
        else:
            shp.line.fill.background()
        tf = shp.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.03 if i == 0 else 0.12); tf.margin_right = Inches(0.05); tf.margin_top = tf.margin_bottom = Inches(0.01)
        small = 6.3 if n <= 7 else 5.8
        third = f"→{hfmt(end)} h · {end / 24:.1f}일" if n <= 7 else f"누적 {hfmt(end)} h"
        lines = ((name, 7.2, True), (f"{hfmt(hh)} h", 11, True), (third, small, False))
        for j, (t, size, bold) in enumerate(lines):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run(); r.text = t; runfmt(r, size, bold, tc)


def calendar(slide, x, y, colw, hdr_h, row_h, year, month, events):
    """events: {date: ([(text, color)], sub, fill)} — 월~일 달력 표."""
    import calendar as _cal
    import datetime as _dt
    first = _dt.date(year, month, 1); start = first - _dt.timedelta(days=first.weekday())
    last = _dt.date(year, month, _cal.monthrange(year, month)[1])
    nw = (last - start).days // 7 + 1
    rows = [["월", "화", "수", "목", "금", "토", "일"]] + [[""] * 7 for _ in range(nw)]
    fills = {}
    for wk in range(nw):
        for d in range(7):
            day = start + _dt.timedelta(days=wk * 7 + d)
            if day in events:
                fills[(wk + 1, d)] = events[day][2]
            elif day.month != month:
                fills[(wk + 1, d)] = "F7F7F7"
    tbl = mk_table(slide, x, y, [colw] * 7, rows, fs=8, fills=fills, left_cols=(), rh_list=[hdr_h] + [row_h] * nw)
    for wk in range(nw):
        for d in range(7):
            day = start + _dt.timedelta(days=wk * 7 + d)
            inm = day.month == month
            dcol = ("C00000" if d == 6 else ("2E75B6" if d == 5 else "595959")) if inm else "BFBFBF"
            paras = [([(str(day.day) if inm else f"{day.month}/{day.day}", 7, inm, dcol)], "l")]
            if day in events:
                main, sub, _ = events[day]
                paras.append(([(t, 10, True, c) for t, c in main], "c"))
                paras.append(([(sub, 6.8, False, "595959")], "c"))
            cell_paras(tbl.cell(wk + 1, d), paras)
    return tbl


KGB = V("KG_B")
T57, TREF = V("T_57"), V("T_REF")

# ============================================================ 1. As-is 공정 흐름 · 공정시간
d = shapes(s1)
set_lines(d["TextBox 182"], ["공용 정제기 1대 · 103℃", "4개 고객 공동 (한솔 '27~)"])
tag_ok = d["Rounded Rectangle 197"]
for nm_, txt in (("Rounded Rectangle 198", "2 h"), ("Rounded Rectangle 199", f"순수 정제 {TREF:.0f} h"), ("Rounded Rectangle 200", "PQC~제품 이송 합계 6 h")):
    copy_fill(tag_ok, d[nm_]); set_lines(d[nm_], [txt])
pos(d["Rounded Rectangle 200"], x=5.06, w=1.70)
delete(d["Rounded Rectangle 201"])
set_lines(d["TextBox 212"], ["5 Gal 글로브 박스 충진", "20 kg/병 · 필터 포함 · 하이닉스·CXMT", "2 h/병 · 9병 18 h (현장 약 2일)"])
set_lines(d["TextBox 216"], ["200 L 수동 충진", "이지켐 140 · 한솔 150 kg/용기 · 필터 포함", "8 h/용기 (이지켐 · 한솔)"])
set_lines(d["TextBox 226"], [["출하 ", "■", " 이지켐  ", "■", " 한솔('27~)"]])
d["TextBox 226"].text_frame.paragraphs[0].runs[3].font.color.rgb = RGBColor.from_string(PUR)
set_lines(d["TextBox 227"], ["OQC·출하 : 5 Gal 9병 2 h · 200 L 1용기 2 h"]); pos(d["TextBox 227"], x=3.90, w=3.10)
widen_flow(s1)
keep_chrome(s1, 3.6)
for sh in list(s1.shapes):
    if sh.name == "TextBox 163":
        delete(sh)
header(s1, 1, "As-is | 공정 흐름과 공정시간", f"현재 정제기 1대 · IQC~FQC {T57:.0f} h (순수 정제 {TREF:.0f} h) · 정제 중 이전 Batch 검사·충진 병행")
sec(s1, 3.76, "① Batch 마일스톤 — 순수 작업시간", "단계별 시간 · 누적")
ms = wb[REF["ms0"][0]]; ms0 = int(REF["ms0"][1][1:])
MS = [[ms.cell(row=ms0 + i, column=c).value for c in range(2, 9)] for i in range(8)]   # 단계, 5Gal h, 시작, 종료, 200L h, 시작, 종료
STY = {"IQC": ("A6A6A6", DARK), "준비·투입": ("F4B183", DARK), "순수 정제": (ORG, "FFFFFF"), "PQC~이송": ("9DC3E6", DARK), "FQC": (BLU, "FFFFFF"),
       "평균 대기": ("F2F2F2", "C00000"), "충진": (GRN, "FFFFFF"), "OQC·출하": ("7F7F7F", "FFFFFF")}


def ms_steps(col, with_wait):
    out, cum = [], 0.0
    for row_ in MS:
        name = row_[0]
        if name.startswith("대기"):
            name = "평균 대기"
        if name == "평균 대기" and not with_wait:
            continue
        hh = row_[col]; cum += hh
        out.append((name, hh, cum, STY[name][0], STY[name][1], name == "평균 대기"))
    return out


def ms_label(slide, y, h, l1, l2, tot):
    tbox(slide, 0.31, y + h / 2 - 0.27, 1.12, 0.56, [[(l1, 8.5, True, DARK)], [(l2, 6.8, False, GRAYT)], [(tot, 10, True, REDC)]])


for i, (l1, l2, col) in enumerate((("5 Gal 9병", "하이닉스 · CXMT", 1), ("200 L 1용기", "이지켐 · 한솔", 4))):
    yy = 4.05 + i * 0.84
    st_ = ms_steps(col, False)
    ms_label(s1, yy, 0.72, l1, l2, f"합계 {hfmt(st_[-1][2])} h")
    chevrons(s1, 1.47, yy, 8.2, 0.72, st_)
sec(s1, 5.76, "② 시간 구분 — 정제기는 47 h만 묶이고, 검사·충진은 다음 Batch 정제와 병행")
rows = [["구분", "5 Gal", "200 L", "내용"],
        ["정제기 점유 (준비·투입 + 순수 정제)", f"{V('T_PREP') + TREF:.0f} h", f"{V('T_PREP') + TREF:.0f} h", "정제기 1대가 묶이는 시간 → Batch 주기를 결정"],
        ["검사 · 이송 (IQC · PQC~이송 · FQC)", "10 h", "10 h", "IQC는 원료 Lot 단위 · PQC~이송은 합계 6 h"],
        ["충진 · OQC", f"{V('fill0') + 2:.0f} h", f"{V('fill1') + 2:.0f} h", "5 Gal 2 h/병 × 9병 + 2 h · 200 L 8 h/용기 + 2 h"]]
mk_table(s1, 0.31, 6.03, [2.75, 0.9, 0.9, 4.82], rows, rh=0.205, fs=8, left_cols=(0, 3), bolds={(1, -1)}, fills={(1, -1): "FFF2E6"})
notes(s1, f"""As-is는 정제기 1대로 하이닉스·CXMT·이지켐(한솔 '27~)을 함께 생산합니다.
57 h는 IQC부터 FQC까지의 현재 보고 기준이고, 순수 정제 {TREF:.0f} h는 57 − IQC 2 − 준비·투입 2 − PQC~이송 6 − FQC 2로 계산한 값입니다(실측 아님). 대화 중 57 − 2 − 6 = 49 h 계산, 과거 약 72 h · 최근 약 53 h 언급이 있어 원자료의 측정 시작·종료점은 확인이 필요합니다.
PQC와 제품 이송은 합쳐서 6 h입니다. 57 h에는 충진·OQC가 포함되지 않습니다.
5 Gal 9병은 2 h × 9 = 18 h 작업이며 현장에서는 근무시간 기준 약 2일로 설명합니다. 200 L은 이지켐·한솔 모두 수동 8 h/용기로 반영했습니다.
경로별 합계 77 h · 67 h는 확인된 시간 항목을 더한 값으로, 달력 기준 납기나 다음 Batch 투입 간격과는 다릅니다. 충진이 진행되는 동안 다음 Batch 정제가 가능합니다.
Batch당 {KGB:.0f} kg과 5 Gal 9병 180 kg의 차이 {KGB - 180:.0f} kg은 잔량·추가 충진·다른 고객 배분 여부를 확인할 항목입니다.""")

# ============================================================ 2. As-is 10월 생산계획 · Batch 완료 주기
cy = wb["03_Batch주기"]
pl0 = int(REF["pl0"][1][1:]); pln = int(REF["pl_n"][1][1:])
PLN = [[cy.cell(row=r_, column=c).value for c in range(2, 9)] for r_ in range(pl0, pln + 1)]   # Batch, 완료일, 요일, 간격일, 간격h, 10월, 누적kg
CYC, WAIT, PLB, PLKG = V("CYC_ASIS"), V("WAIT_ASIS"), V("PL_B"), V("PL_KG")
D3, D2 = V("PL_D3"), V("PL_D2")
sP = new_slide()
header(sP, 2, "As-is | 10월 생산계획 — Batch 완료 주기",
       f"10월 {PLB:.0f} Batch ({PLN[1][0]}~{PLN[-1][0]}) = {f0(PLKG)} kg · 평균 완료 간격 {CYC / 24:.1f}일 = {CYC:.1f} h")
sec(sP, 1.05, "① 2026년 10월 생산계획", "완료일 = 생산 및 충진 완료", w=5.9)
ev = {}
for b_, d_, _, _, _, oct_, kg_ in PLN:
    dd = d_.date()
    ev[dd] = ([(b_, ORG if oct_ else GRAYT)], f"누적 {f0(kg_)} kg" if oct_ else "9월 완료", "FFF2E6" if oct_ else "F2F2F2")
calendar(sP, 0.31, 1.33, 0.84, 0.26, 0.6, 2026, 10, ev)
sec(sP, 1.05, "② 완료 간격", x=6.42, w=3.2)
rows = [["구분", "횟수", "시간"], ["3일 간격", f"{D3:.0f}회", "72 h"], ["2일 간격", f"{D2:.0f}회", "48 h"],
        [f"평균 ({PLN[0][0]}→{PLN[-1][0]})", f"{(PLN[-1][1] - PLN[0][1]).days}일 ÷ {len(PLN) - 1}", f"{CYC:.1f} h"],
        ["10월 완료", f"{PLB:.0f} Batch", f"{f0(PLKG)} kg"]]
mk_table(sP, 6.42, 1.33, [1.4, 0.9, 0.97], rows, rh=0.33, fs=8.2, bolds={(3, -1)}, fills={(3, -1): "FFF2E6"})
sec(sP, 3.15, "③ 정제기 주기 구성", x=6.42, w=3.2)
rows = [["구분", "시간", "비율"], ["순수 (준비·투입 + 정제)", f"{V('T_PREP') + TREF:.0f} h", f"{(V('T_PREP') + TREF) / CYC * 100:.0f}%"],
        ["대기 · 전환", f"{WAIT:.1f} h", f"{WAIT / CYC * 100:.0f}%"], ["완료 간격", f"{CYC:.1f} h", "100%"]]
mk_table(sP, 6.42, 3.43, [1.6, 0.85, 0.82], rows, rh=0.315, fs=8.2, bolds={(2, -1), (3, -1)}, colors={(2, 1): REDC, (2, 2): REDC})
sec(sP, 4.85, "④ 월 생산량 비교 — 정제기 1대")
c0 = int(REF["cmp0"][1][1:])
CMP = [[cy.cell(row=c0 + i, column=c).value for c in range(2, 10)] for i in range(6)]   # 기준, 대, Batch, kg, 연간, 주기, 순수, 대기
rows = [["기준", "Batch/월", "kg/월", "연간 (t)", "완료 간격 (h)", "산출 근거"]]
for i, basis in ((0, "720 h ÷ 45 h (점유·전환·대기 제외)"), (1, f"{(PLN[-1][1] - PLN[0][1]).days}일 × 24 h ÷ {len(PLN) - 1} Batch"),
                 (2, f"협의 계획 12 Batch × {KGB:.0f} kg"), (4, "26,200 kg ÷ 12")):
    c_ = CMP[i]
    rows.append([c_[0], f"{c_[2]:.1f}" if i == 4 else f"{c_[2]:.0f}", f0(c_[3]), f"{c_[4]:.1f}", f"{c_[5]:.1f}", basis])
mk_table(sP, 0.31, 5.13, [2.55, 0.9, 1.0, 0.95, 1.2, 2.77], rows, rh=0.33, fs=8.3, left_cols=(0, 5), bolds={(2, -1), (3, -1)},
         fills={(2, -1): "FFF2E6", (3, -1): "EAF1FB"}, colors={(3, 4): REDC})
notes(sP, f"""10월 생산계획의 표시일은 해당 Batch의 생산과 충진이 모두 끝나는 완료 예정일이며, Batch 번호는 각각 독립된 생산 Batch입니다.
10월에는 {PLN[1][0]}~{PLN[-1][0]} {PLB:.0f} Batch가 완료되어 {f0(PLKG)} kg입니다. 9월 마지막 {PLN[0][0]}(9/28)부터 {PLN[-1][0]}(10/29)까지 {(PLN[-1][1] - PLN[0][1]).days}일 동안 {len(PLN) - 1}번 완료되어 평균 간격은 {CYC / 24:.2f}일 = {CYC:.1f} h입니다(3일 간격 {D3:.0f}회, 2일 간격 {D2:.0f}회).
정제기는 Batch마다 준비·투입 2 h + 순수 정제 {TREF:.0f} h = {V('T_PREP') + TREF:.0f} h를 순수하게 사용하고, 나머지 {WAIT:.1f} h는 제품 이송 중 점유·세척·전환·투입 대기 등입니다.
순수 정제만으로 계산한 720 ÷ 45 = 16 Batch는 이론값이고, 실제 계획은 대기를 포함해 11 Batch입니다. 현재 연간 Capa. 26.2톤은 월 11.5 Batch 상당입니다.""")

# ============================================================ 3. As-is 순수 작업 vs 대기
sW = new_slide()
W = [[cy.cell(row=int(REF[f"w{i}"][1][1:]), column=c).value for c in range(2, 9)] for i in range(4)]   # 구분, 순수, 대기, 실제, 일, 비율, 근거
header(sW, 3, "As-is | Batch 소요시간 — 순수 작업 vs 대기",
       f"정제기 주기 {CYC:.1f} h = 순수 {W[0][1]:.0f} h + 대기 {WAIT:.1f} h · 5 Gal 1 Batch 실제 약 {W[1][3]:.0f} h ({W[1][4]:.1f}일)")
sec(sW, 1.05, "① 순수 작업 vs 대기 (평균)", "10월 생산계획 완료 간격 기준")
rows = [["구분", "순수 작업", "대기", "실제 소요", "일 환산", "순수 비율", "산출 근거"]]
for w_ in W:
    rows.append([w_[0], f"{hfmt(w_[1])} h", f"{hfmt(w_[2])} h", f"{hfmt(w_[3])} h", f"{w_[4]:.1f}일", f"{w_[5] * 100:.0f}%", w_[6]])
mk_table(sW, 0.31, 1.33, [2.35, 0.9, 0.85, 1.0, 0.8, 0.85, 2.62], rows, rh=0.32, fs=8.2, left_cols=(0, 6),
         bolds={(1, 3), (2, 3), (3, 3), (1, 2), (2, 2), (3, 2), (4, 2)}, colors={(i, 2): REDC for i in range(1, 5)}, fills={(1, -1): "FFF2E6"})
sec(sW, 3.12, "② 평균 Batch 마일스톤 — 대기 포함", "단계 시간 · 누적 시간 (일)")
for i, (l1, l2, col) in enumerate((("5 Gal 9병", "하이닉스 · CXMT", 1), ("200 L 1용기", "이지켐 · 한솔", 4))):
    yy = 3.42 + i * 0.86
    st_ = ms_steps(col, True)
    ms_label(sW, yy, 0.74, l1, l2, f"{hfmt(st_[-1][2])} h")
    chevrons(sW, 1.47, yy, 8.2, 0.74, st_)
sec(sW, 5.25, "③ 대기 구성 — 실측으로 확인할 항목")
rows = [["항목", "내용", "확인"],
        ["제품 이송 중 정제기 점유", "PQC~이송 6 h 중 정제기가 묶이는 시간", "실측"],
        ["세척 · 전환", "Batch 사이 정제기 세척 · 전환", "실측"],
        ["투입 대기", "원료 IQC 승인 · Mix·Premix 준비 · 작업 인원", "실측"],
        ["충진 대기 (근무시간)", f"5 Gal 18 h 작업 → 약 2일 경과 (대기 약 {W[3][2]:.0f} h)", "교대 조건"]]
mk_table(sW, 0.31, 5.52, [2.35, 5.2, 1.82], rows, rh=0.26, fs=8, left_cols=(0, 1))
notes(sW, f"""우리가 계산한 77 h(5 Gal)·67 h(200 L)는 각 공정의 순수 작업시간을 더한 값이라 대기시간이 들어 있지 않습니다.
10월 생산계획의 평균 완료 간격 {CYC:.1f} h에서 정제기 순수 점유 {W[0][1]:.0f} h(준비·투입 2 + 순수 정제 {TREF:.0f})를 빼면 Batch마다 평균 {WAIT:.1f} h의 대기·전환이 있습니다(순수 비율 {W[0][5] * 100:.0f}%).
이 대기를 더하면 1 Batch의 실제 소요는 5 Gal 약 {W[1][3]:.1f} h({W[1][4]:.1f}일), 200 L 약 {W[2][3]:.1f} h({W[2][4]:.1f}일)입니다. 완료 간격({CYC:.1f} h)보다 길기 때문에 앞 Batch를 충진하는 동안 다음 Batch를 정제하는 병행 운전이 전제입니다.
마일스톤의 대기 위치는 평균값을 FQC 뒤에 모아 표시한 것으로, 실제로는 제품 이송 점유·세척·전환·투입 대기·충진 대기(근무시간)로 나뉩니다. 5 Gal 충진은 18 h 작업이지만 근무시간 기준 약 2일이 걸려 충진 단계만 봐도 대기가 약 {W[3][2]:.0f} h입니다.
협의 계획 기준으로 정제기 1대 월 12 Batch는 대기를 {WAIT:.1f} h에서 {CMP[2][7]:.1f} h로 줄여야 하고, 정제기 2대 월 20 Batch는 설비별 대기 {V('WAIT_RF'):.1f} h로 현재보다 여유가 있습니다(6장).""")

# ============================================================ 3. To-be 공정 흐름 · 일정
d = shapes(s2)
vessel = [sh for sh in s2.shapes if sh.name.startswith("Round Same Side Corner")][0]
noz, vtxt = d["Rectangle 177"], d["TextBox 179"]
for sh in (vessel, noz, vtxt):
    sh.left = Emu(int(sh.left - 0.36 * E))
set_lines(vtxt, ["정제기1"]); vtxt.text_frame.paragraphs[0].runs[0].font.size = Pt(5.5)
for sh in (vessel, noz, vtxt):
    c = clone(s2, sh); c.left = Emu(int(sh.left + 0.70 * E))
    if c.has_text_frame and c.text_frame.text:
        set_lines(c, ["정제기2"])
tbox(s2, 3.93, 1.40, 0.30, 0.14, [("시차", 5.5, True, REDC)], align="c")
pos(d["Connector 193"], w=3.48 - 2.86)
c194 = d["Connector 194"]; old_end = (c194.left + c194.width) / E; pos(c194, x=4.60, w=old_end - 4.60)
set_lines(d["TextBox 182"], ["정제기 1·2 · 투입 시점 엇갈림", "4개 고객 공용 배정"])
tag_ok = d["Rounded Rectangle 197"]
for nm_, txt in (("Rounded Rectangle 198", "2 h"), ("Rounded Rectangle 199", f"순수 정제 {TREF:.0f} h (설비별)"), ("Rounded Rectangle 200", "PQC~제품 이송 합계 6 h")):
    copy_fill(tag_ok, d[nm_]); set_lines(d[nm_], [txt])
pos(d["Rounded Rectangle 199"], x=3.38, w=1.34); pos(d["Rounded Rectangle 200"], x=5.06, w=1.70)
delete(d["Rounded Rectangle 201"])
set_lines(d["TextBox 212"], ["5 Gal 글로브 박스 충진 (유지)", "20 kg/병 · 필터 포함 · 하이닉스·CXMT", "2 h/병 · 9병 18 h"])
set_lines(d["TextBox 216"], ["200 L 수동 → ARS ('27.7~)", "이지켐 140 · 한솔 150 kg/용기 · 필터 포함", "수동 8 h/용기 · ARS 단계별 시간 실측"])
set_lines(d["TextBox 226"], [["출하 ", "■", " 이지켐  ", "■", " 한솔"]])
d["TextBox 226"].text_frame.paragraphs[0].runs[3].font.color.rgb = RGBColor.from_string(PUR)
set_lines(d["TextBox 227"], ["OQC·출하 : 5 Gal 9병 2 h · 200 L 1용기 2 h"]); pos(d["TextBox 227"], x=3.90, w=3.10)
widen_flow(s2)
keep_chrome(s2, 3.6)
for sh in list(s2.shapes):
    if sh.name == "TextBox 163":
        delete(sh)
header(s2, 5, "To-be | 정제기 2대 운영과 투자 일정", "정제기 2대의 투입 시점을 엇갈리게 운영하여 생산을 병행하고, 후공정 부하를 분산하는 계획")
sec(s2, 3.76, "① Case별 리플럭스 도입 · 월 Batch 계획", f"1대 월 {V('MB_ONE'):.0f} Batch · 2대 월 {V('MB_TWO'):.0f} Batch · {KGB:.0f} kg/Batch")
M16 = [(2026, m) for m in (9, 10, 11, 12)] + [(2027, m) for m in range(1, 13)]
yrs = ["구분", "2026년", "", "", "", "2027년"] + [""] * 11
mons = [""] + [f"{m}월" for _, m in M16]
c1 = wb["05_Case1"]; c3 = wb["07_Case3"]; P1 = MAP["CROW"]["I"]; P3 = MAP["CROW"]["III"]
bI = [c1.cell(row=P1["batch"], column=3 + i).value for i in range(16)]
bIII = [c3.cell(row=P3["batch"], column=3 + i).value for i in range(16)]
def bcell(v, i, intro):
    if i == 0:
        return "-"
    return f"{v} 도입" if i == intro else (f"{v} 운전" if i == intro + 1 else str(v))
r_a = ["Case I · II  Batch"] + [bcell(v, i, 4) for i, v in enumerate(bI)]
r_b = ["Case III  Batch"] + [bcell(v, i, 7) for i, v in enumerate(bIII)]
r_ars = ["200 L 충진"] + ["수동 8 h/용기"] + [""] * 9 + ["ARS '27.7~ (이지켐 · 추가)"] + [""] * 5
r_ars = r_ars[:17]
rows = [yrs, mons, r_a, r_b, r_ars]
fl = {(0, -1): "595959", (1, -1): "F2F2F2"}
for ri, arr, intro in ((2, bI, 4), (3, bIII, 7)):
    for i, v in enumerate(arr):
        j = i + 1
        if i == 0:
            fl[(ri, j)] = "FFFFFF"
        elif v == 6:
            fl[(ri, j)] = "D9D9D9"
        elif i > intro:
            fl[(ri, j)] = "FDE9E7"
        else:
            fl[(ri, j)] = "F2F2F2"
fl.update({(4, 11): "E4DFEC"})
cl = {(1, j): DARK for j in range(17)}
cl.update({(2, 5): REDC, (2, 6): REDC, (3, 8): REDC, (3, 9): REDC, (4, 1): GRAYT, (4, 11): PUR})
mk_table(s2, 0.31, 4.02, [1.45] + [0.495] * 16, rows, rh=0.27, fs=7.5,
         merges=[(0, 1, 4), (0, 5, 16), (4, 1, 10), (4, 11, 16)],
         fills=fl, colors=cl, bolds={(2, -1), (3, -1), (1, -1)}, hdr_rows=1)
sec(s2, 5.5, "② 생산능력 기준 (월 Batch × 198 kg)")
rows = [["구분", "월 Batch", "kg/월", "연간 환산", "비고"],
        ["정제기 1대 (현재)", f"{V('MB_ONE'):.0f}", f0(V("MB_ONE") * KGB), f"{V('MB_ONE') * KGB * 12 / 1000:.1f}톤", "현재 Capa. 26.2톤 수준"],
        ["정제기 2대 (리플럭스 후)", f"{V('MB_TWO'):.0f}", f0(V("MB_TWO") * KGB), f"{V('CAPA_TWO'):.1f}톤", "리플럭스 47.2톤 계획 수준"],
        ["대정비 월 ('27.4)", f"{V('MB_MAINT'):.0f}", f0(V("MB_MAINT") * KGB), "-", "Case I·II · III 공통"]]
mk_table(s2, 0.31, 5.78, [2.6, 1.0, 1.2, 1.3, 3.27], rows, rh=0.27, fs=8.3, left_cols=(0, 4), bolds={(2, -1)}, fills={(2, -1): "FFF2E6"})
notes(s2, f"""리플럭스 이후에는 정제기 2대를 활용합니다. 정제기 1을 먼저 투입·운전하고, 초류 진행에 맞춰 정제기 2를 투입해 두 설비의 운전 구간을 겹칩니다(투입 간격은 미확정). 정제기별 Batch 시간이 절반으로 줄어드는 것이 아니라 두 설비의 생산을 겹쳐 전체 생산량을 늘립니다.
생산팀 협의 기준: Batch size 200 kg × Yield 99% = {KGB:.0f} kg/Batch, Batch time 55 h, 원단위 1.010. 정제기 1대 월 {V('MB_ONE'):.0f} Batch({f0(V('MB_ONE') * KGB)} kg), 2대 월 {V('MB_TWO'):.0f} Batch({f0(V('MB_TWO') * KGB)} kg), 4월 대정비 {V('MB_MAINT'):.0f} Batch({f0(V('MB_MAINT') * KGB)} kg).
Case I·II: '27.1 리플럭스 도입 운영, '27.2 운전(16 Batch), 3월부터 20 Batch. Case III: '27.4 도입, '27.5부터 20 Batch — 1~3월은 1대 12 Batch로 유지합니다.
정제기 2대 월 20 Batch는 연 {V('CAPA_TWO'):.1f}톤 수준으로 리플럭스 47.2톤 계획과 맞습니다. 47.2톤이 현재의 2배(52.4톤)가 아닌 것은 Mix·Premix 준비 등 설비 시간이 필요하기 때문입니다.
200 L은 '27.7부터 ARS(이지켐·추가 물량), 5 Gal은 글로브 박스를 유지합니다. 추가 정제기의 개조/신규 여부와 설비 번호(2070·2060 등)는 도면과 대조해 확인합니다.""")

# ============================================================ 5. To-be 운영 마일스톤
sB = new_slide()
IRF, WRF, OFF = V("INT_RF"), V("WAIT_RF"), V("OFFSET_H")
tb, tbm = V("tobe_b"), V("tobe_bm")
header(sB, 6, "To-be | 정제기 2대 시차 운전 — 운영 마일스톤",
       f"정제기 2대 월 {tbm:.0f} Batch = {f0(tbm * KGB)} kg · 설비별 주기 {IRF:.0f} h = 순수 47 h + 대기 {WRF:.0f} h")
tb0 = int(REF["tb0"][1][1:])
comp = []
for i in range(16):
    for u, c0_ in ((0, 3), (1, 7)):
        h_ = cy.cell(row=tb0 + i, column=c0_ + 1).value; d_ = cy.cell(row=tb0 + i, column=c0_ + 2).value
        if cy.cell(row=tb0 + i, column=c0_ + 3).value == 1:
            comp.append((h_, u, d_.date()))
comp.sort()
tm = cy.cell(row=int(REF["INT_RF"][1][1:]) + 3, column=3).value.date()
ev, cnt = {}, 0
for h_, u, d_ in comp:
    cnt += 1
    runs, _, _ = ev.get(d_, ([], "", ""))
    runs = runs + [(("①" if u == 0 else "②") + " ", ORG if u == 0 else BLU)]
    ev[d_] = (runs, f"누적 {cnt} Batch", "FFF2E6")
sec(sB, 1.05, f"① {tm.year}년 {tm.month}월 운영 예시 — Batch 완료일", f"① 정제기 1 · ② 정제기 2 (시차 {OFF:.0f} h 가정)", w=5.9)
calendar(sB, 0.31, 1.33, 0.84, 0.26, 0.6, tm.year, tm.month, ev)
n1 = sum(1 for c in comp if c[1] == 0); n2 = len(comp) - n1
sec(sB, 1.05, "② 운영 결과", x=6.42, w=3.2)
rows = [["구분", f"{tm.month}월"], ["정제기 ① 완료", f"{n1} Batch"], ["정제기 ② 완료", f"{n2} Batch"], [f"{tm.month}월 합계", f"{tb:.0f} Batch"],
        ["생산량", f"{f0(tb * KGB)} kg"], ["협의 계획 (월)", f"{tbm:.0f} Batch"], ["협의 계획 kg", f"{f0(tbm * KGB)} kg"]]
mk_table(sB, 6.42, 1.33, [1.75, 1.52], rows, rh=0.355, fs=8.2, bolds={(3, -1), (4, -1)}, fills={(3, -1): "FFF2E6", (4, -1): "FFF2E6"})
sec(sB, 4.85, "③ 월 생산량 비교 — 설비별 주기 = 순수 47 h + 대기")
rows = [["기준", "정제기", "Batch/월", "kg/월", "설비별 주기 (h)", "순수 (h)", "대기 (h)"]]
for i in (1, 2, 3):
    c_ = CMP[i]
    rows.append([c_[0], f"{c_[1]:.0f}대", f"{c_[2]:.0f}", f0(c_[3]), f"{c_[5]:.1f}", f"{c_[6]:.0f}", f"{c_[7]:.1f}"])
mk_table(sB, 0.31, 5.13, [2.6, 0.8, 0.95, 1.05, 1.35, 0.95, 1.67], rows, rh=0.36, fs=8.5, bolds={(3, -1)},
         fills={(3, -1): "FFF2E6"}, colors={(2, 6): REDC, (1, 6): REDC, (3, 6): "385723"})
notes(sB, f"""정제기 2대 운영 예시입니다. 협의 계획 월 {tbm:.0f} Batch를 평균 월 730 h에 2대로 나누면 설비별 주기는 {IRF:.1f} h입니다. 정제기 2는 정제기 1보다 {OFF:.0f} h 늦게 투입하는 것으로 가정했습니다(실제 시차는 초류 진행을 보고 결정).
달력의 완료일은 투입 후 5 Gal 기준 실제 소요(순수 77 h + 대기 {WAIT:.1f} h = {V('LT_5G'):.1f} h)가 지난 날짜입니다. {tm.year}년 {tm.month}월에는 정제기 ① {n1} Batch, ② {n2} Batch, 합계 {tb:.0f} Batch = {f0(tb * KGB)} kg이 완료됩니다.
설비별 주기 {IRF:.1f} h = 순수 47 h + 대기 {WRF:.1f} h로, 10월 생산계획의 대기 {WAIT:.1f} h보다 여유가 있어 2대 월 20 Batch는 현재 운영 수준으로 달성 가능합니다.
반대로 정제기 1대 월 12 Batch(2026.10~12, Case III는 2027.1~3까지)는 주기 {CMP[2][5]:.1f} h, 대기 {CMP[2][7]:.1f} h가 필요해 현재 10월 계획(11 Batch · 대기 {WAIT:.1f} h)보다 대기를 {WAIT - CMP[2][7]:.1f} h 줄여야 합니다.""")

# ============================================================ 6. Case 비교
sV = new_slide()
vw = wb["08_Case비교"]
CV = [[vw.cell(row=6 + i, column=c).value for c in range(2, 12)] for i in range(3)]   # Case, 리플럭스, Batch, 생산, 판매, 차이, 말재고, 최저, 시점, 최저재고일
header(sV, 7, "Case 비교 — 2027년 생산 · 판매 · 재고",
       f"'27 말 재고  Case I {f0(CV[0][6])} · Case II {f0(CV[1][6])} · Case III {f0(CV[2][6])} kg (4월 {f0(CV[2][7])} kg)")
sec(sV, 1.05, "① 2027년 연간 비교 (kg)", "재고 = 전월 재고 + 생산 − 판매 · 26.9말 2,300 kg 시작")
SALE = {"I": "계획 比 CXMT·이지켐 조정", "II": "추가 물량 (월 300 kg)", "III": "추가 물량 (월 300 kg)"}
rows = [["Case", "판매", "리플럭스", "생산", "판매", "생산 − 판매", "'27 말 재고", "최저 재고 (시점)", "최저 재고일"]]
for k, v_ in zip(("I", "II", "III"), CV):
    rows.append([v_[0], SALE[k], v_[1], f0(v_[3]), f0(v_[4]), sg(v_[5]), f0(v_[6]), f"{f0(v_[7])} ('{v_[8]:%y.%-m})", f"{v_[9]:.1f}개월"])
cl = {(i, 5): ("385723" if rows[i][5].startswith("+") else REDC) for i in (1, 2, 3)}
cl.update({(3, 7): REDC, (3, 8): REDC})
mk_table(sV, 0.31, 1.33, [0.75, 1.9, 1.55, 0.8, 0.8, 0.85, 0.85, 1.07, 0.8], rows, rh=0.36, fs=8, left_cols=(0, 1, 2),
         bolds={(i, 6) for i in (1, 2, 3)} | {(i, 0) for i in (1, 2, 3)}, colors=cl, fills={(3, -1): "FFF2F2"})
sec(sV, 2.95, "② 월별 재고 (kg) · 재고일 (총판매량 기준, 개월)")
rows = [["구분"] + [f"{m}월" for _, m in M16]]
mc = int(REF["cmpm"][1][1:])
cl, fl = {}, {}
for i in range(6):
    vals = [vw.cell(row=mc + i, column=3 + k).value for k in range(16)]
    lab = vw.cell(row=mc + i, column=2).value
    if i % 2 == 0:
        rows.append([lab] + [f0(v) for v in vals])
        for k, v in enumerate(vals):
            if v < 500:
                cl[(len(rows) - 1, k + 1)] = REDC
        fl[(len(rows) - 1, -1)] = "FFF2E6"
    else:
        rows.append([lab] + [f"{v:.1f}" for v in vals])
        for k, v in enumerate(vals):
            if v < 0.5:
                cl[(len(rows) - 1, k + 1)] = REDC
mk_table(sV, 0.31, 3.23, [1.35] + [0.5] * 16, rows, rh=0.33, fs=7.6, colors=cl, fills=fl, bolds={(1, -1), (3, -1), (5, -1)})
sec(sV, 5.7, "③ 판단")
rows = [["Case", "요약"],
        ["Case I", "판매 조정 · 1월 도입 → 재고 꾸준히 증가 ('27 말 2.2개월) — 재고 과다 관리 필요"],
        ["Case II", "추가 물량 · 1월 도입 → 4월 대정비 때 2,420 kg (0.9개월)까지 감소 후 1개월 수준 유지"],
        ["Case III", "추가 물량 · 4월 도입 → 4월 재고 44 kg (0.0개월) — 외부 상품 도입 없이는 공급 부족"]]
mk_table(sV, 0.31, 5.98, [1.0, 8.37], rows, rh=0.215, fs=8, left_cols=(0, 1), colors={(3, 1): REDC})
notes(sV, f"""오늘 생산팀과 협의한 세 가지 Case입니다. 공통 조건: 26.9말 재고 2,300 kg, 2026.10~12 정제기 1대 월 12 Batch(2,376 kg), {KGB:.0f} kg/Batch, 4월 대정비 6 Batch(1,188 kg), 2대 운전 후 월 20 Batch(3,960 kg).
Case I은 판매를 계획 대비 CXMT·이지켐 물량으로 조정(2027 판매 {f0(CV[0][4])} kg)하고 1월 리플럭스 도입·2월 운전입니다. 2027 생산 {f0(CV[0][3])} kg으로 재고가 '27 말 {f0(CV[0][6])} kg까지 늘어납니다.
Case II는 월 300 kg 추가 물량(2027 판매 {f0(CV[1][4])} kg)에 1월 도입으로, 4월 대정비 때 재고가 {f0(CV[1][7])} kg까지 줄었다가 '27 말 {f0(CV[1][6])} kg입니다.
Case III은 추가 물량에 리플럭스를 4월 도입·5월 운전으로 늦춘 경우로, 1~3월 1대 12 Batch로는 판매를 따라가지 못해 4월 재고가 {f0(CV[2][7])} kg(0.0개월)까지 떨어집니다. 외부 상품 도입을 반영해야 합니다.
재고일(개월)은 협의 자료 값을 그대로 옮겼습니다(산식 확인 필요).""")

# ============================================================ 7~9. Case 상세
CASE_T = {"I": ("판매 : 계획 比 CXMT, 이지켐 물량 조정", "생산 : 1월 Reflux Column 도입 운영 (2월 운전)", ""),
          "II": ("판매 : 추가 물량", "생산 : 1월 Reflux Column 도입 운영 (2월 운전)", ""),
          "III": ("판매 : 추가 물량", "생산 : 4월 Reflux Column 도입 운영 (5월 운전)", "외부 상품 도입 반영 필요")}
case_slides = []
for num, (k, sheet) in enumerate((("I", "05_Case1"), ("II", "06_Case2"), ("III", "07_Case3")), start=8):
    s = new_slide(); case_slides.append(s)
    cw_ = wb[sheet]; P = MAP["CROW"][k]
    g = {key: [cw_.cell(row=P[key], column=3 + i).value for i in range(17)] for key in
         ("batch", "prod", "sale", "skh", "cx", "ez", "add", "inv", "d_tot", "d_sk", "d_skc")}
    v_ = CV[("I", "II", "III").index(k)]
    t1, t2, t3 = CASE_T[k]
    header(s, num, f"Case {k} | 2026.9 ~ 2027.12 생산 · 판매 · 재고",
           f"{t1[5:]} · {t2[5:].replace(' Column', '')}" + (f" → {t3}" if t3 else ""))
    sec(s, 1.05, "① 월별 생산 · 판매 · 재고 (kg)", f"Batch × {KGB:.0f} kg · 재고 = 전월 재고 + 생산 − 판매")
    yrow = ["구분", "2026년", "", "", "", "2027년"] + [""] * 11 + ["27년 합계"]
    mrow_ = [""] + [f"{m}월" for _, m in M16] + [""]
    rows = [yrow, mrow_]
    def fmt_row(key, fm):
        vals = g[key]
        out = []
        for i, v in enumerate(vals):
            if i == 16 and key.startswith("d_"):
                out.append("")
            elif v in (None, ""):
                out.append("")
            elif isinstance(v, str):
                out.append(v)
            else:
                out.append(fm(v))
        return out
    spec = (("batch", "Batch (월)", lambda v: f"{v:.0f}"), ("prod", "생산", f0), ("sale", "판매", f0), ("skh", "SKH", f0), ("cx", "CXMT", f0),
            ("ez", "이지켐", f0), ("add", "추가", f0), ("inv", "재고", f0), ("d_tot", "재고일 (총판매량)", lambda v: f"{v:.1f}"),
            ("d_sk", "재고일 (SKHY 기준)", lambda v: f"{v:.1f}"), ("d_skc", "재고일 (SKHY·CXMT)", lambda v: f"{v:.1f}"))
    for key, lab, fm in spec:
        rows.append([lab] + fmt_row(key, fm))
    fl = {(0, -1): "595959", (1, -1): "595959", (9, -1): "FFF2E6"}
    cl = {}
    intro = 4 if k != "III" else 7
    for i in range(16):
        b = g["batch"][i]
        fl[(3, i + 1)] = "FFFFFF" if i == 0 else ("D9D9D9" if b == 6 else ("FDE9E7" if i > intro else "F2F2F2"))
        if isinstance(g["inv"][i], (int, float)) and g["inv"][i] < 500:
            cl[(9, i + 1)] = REDC
        for ri, key in ((10, "d_tot"), (11, "d_sk"), (12, "d_skc")):
            if isinstance(g[key][i], (int, float)) and g[key][i] < 0.5:
                cl[(ri, i + 1)] = REDC
    for ri in range(5, 9):
        for j in range(1, 18):
            cl.setdefault((ri, j), "595959")
    mk_table(s, 0.31, 1.33, [1.3] + [0.47] * 16 + [0.55], rows, rh=0.315, fs=7.6, merges=[(0, 1, 4), (0, 5, 16)],
             fills=fl, colors=cl, bolds={(2, -1), (3, -1), (9, -1)}, hdr_rows=2)
    sec(s, 5.55, "② 2027년 요약")
    rows = [["2027 Batch", "2027 생산", "2027 판매", "생산 − 판매", "'27 말 재고", "최저 재고 (시점)", "최저 재고일"],
            [f"{v_[2]:.0f}", f"{f0(v_[3])} kg", f"{f0(v_[4])} kg", f"{sg(v_[5])} kg", f"{f0(v_[6])} kg", f"{f0(v_[7])} kg ('{v_[8]:%y.%-m})", f"{v_[9]:.1f}개월"]]
    mk_table(s, 0.31, 5.83, [1.2, 1.3, 1.3, 1.3, 1.3, 1.67, 1.3], rows, rh=0.36, fs=9, left_cols=(),
             bolds={(1, -1)}, colors={(1, 3): "385723" if v_[5] >= 0 else REDC, (1, 5): REDC if v_[7] < 500 else DARK})
    notes(s, f"""Case {k}. {t1} / {t2}{(' / ' + t3) if t3 else ''}.
생산은 월 Batch × {KGB:.0f} kg(Batch size 200 kg × Yield 99%)이며, 2026.9월 생산 1,770 kg은 협의 자료 값입니다. 정제기 1대 기간은 월 12 Batch, 리플럭스 운전 후 월 20 Batch, 4월은 대정비로 6 Batch입니다.
2027 생산 {f0(v_[3])} kg, 판매 {f0(v_[4])} kg으로 {sg(v_[5])} kg, 재고는 '27 말 {f0(v_[6])} kg이고 최저는 {v_[8]:%Y.%-m}월 {f0(v_[7])} kg입니다.
판매는 SKH 연 22,720 kg, CXMT 4월부터(580 → 7월 780 kg), 이지켐 560 → 6월부터 840 kg{', 추가 물량 월 300 kg' if k != 'I' else ''}입니다. 2026년 판매 합계는 협의 자료 값(10·12월은 고객별 합보다 140 kg 큼)을 그대로 썼습니다.
재고일(개월)은 협의 자료 값입니다(산식 확인 필요).""")

# ============================================================ 10. 후공정 부하
sL = new_slide()
l27 = {k: lrow(2027, k)[0] for k in ("g5bt", "g5h", "gbav", "gbld", "ezc", "hsc", "l2h", "g5oqc", "l2oqc", "qch", "nb")}
peak = max(l27["gbld"])
header(sL, 11, "후공정 부하 검토 — 충진 · 검사 · OQC", f"정제능력과 출하능력은 별개 · 5 Gal 충진 최대 월 {max(l27['g5h']):.0f} h = 1교대 {l27['gbav'][0]:.0f} h의 {peak * 100:.0f}%")
sec(sL, 1.05, "① 월 작업량 (Case II 판매 기준)", "추가 물량 포함 — 판매 최대 Case")
cols = (("'27.1", l27, 0), ("'27.4", l27, 3), ("'27.7", l27, 6), ("'27.12", l27, 11))
items = (("5 Gal 병", "g5bt", f0, "하이닉스 + CXMT ÷ 20 kg"), ("5 Gal 충진 h", "g5h", f0, "2 h/병"),
         ("글로브 박스 가용 h", "gbav", f0, "1교대 8 h × 22일"), ("글로브 박스 부하율", "gbld", lambda v: f"{v * 100:.0f}%", "1교대 기준"),
         ("200 L 용기", None, f0, "이지켐 140 · 추가 150 kg"), ("200 L 충진 h", "l2h", f0, "8 h/용기"),
         ("OQC·출하 h", None, f0, "9병 2 h · 1용기 2 h"), ("검사 h (PQC~FQC)", "qch", f0, "8 h × 필요 Batch"),
         ("필요 Batch", "nb", f1, "판매 ÷ 198"))
rows = [["항목"] + [c[0] for c in cols] + ["기준"]]
for lab, key, fm, basis in items:
    row_ = [lab]
    for _, L_, i in cols:
        if key is None and "용기" in lab:
            v = L_["ezc"][i] + L_["hsc"][i]
        elif key is None:
            v = L_["g5oqc"][i] + L_["l2oqc"][i]
        else:
            v = L_[key][i]
        row_.append(fm(v))
    rows.append(row_ + [basis])
cl = {(4, j): REDC for j in range(1, 5) if float(rows[4][j].rstrip("%")) > 100}
mk_table(sL, 0.31, 1.33, [2.3, 1.05, 1.05, 1.05, 1.05, 2.87], rows, rh=0.28, fs=8, left_cols=(0, 5), bolds={(2, -1), (4, -1)}, colors=cl)
sec(sL, 4.30, "② 병목 검토")
rows = [["자원", "현재", "리플럭스 이후", "확인 필요"],
        ["5 Gal 글로브 박스", "9병 18 h (약 2일)", f"월 {max(l27['g5h']):.0f} h > 1교대 176 h", "교대 · 인원 · 글로브 박스 수"],
        ["200 L 충진", "수동 8 h/용기", "ARS '27.7~ (시간 비슷할 수 있음)", "ARS 단계별 시간 실측"],
        ["검사 (PQC · FQC)", "8 h/Batch", "Batch 증가 · 회수 시점 겹침", "검사 인력 · 승인 대기"],
        ["Product Tank", "PQC 합격 후 이송", "정제기 2대 제품 보관", "Tank 수 · 용량"],
        ["충진 · 포장", "1 Batch 약 2일", "목표 1~1.5일", "개선 방안 · 일정"]]
mk_table(sL, 0.31, 4.58, [1.9, 2.0, 2.7, 2.77], rows, rh=0.35, fs=8, left_cols=(0, 1, 2, 3))
notes(sL, f"""후공정 작업량은 판매 물량이 가장 큰 Case II로 계산했습니다. 5 Gal은 (SKH + CXMT) ÷ 20 kg × 2 h, 200 L은 이지켐 140 kg · 추가 물량 150 kg 용기(한솔 기준 가정) × 8 h(수동 기준)입니다.
2027년 하반기에는 5 Gal 충진이 월 약 {max(l27['g5h']):.0f} h로 1교대(8 h × 22일 = 176 h)를 넘습니다. 정제기가 2대가 되어도 충진·검사·Tank가 따라오지 못하면 출하량은 늘지 않으므로 교대·인원 계획이 함께 필요합니다.
ARS는 실제 충진은 더 길 수 있지만 용기 투입·반출·퍼지를 포함하면 수동과 전체 시간이 비슷할 수 있다는 설명이 있어, 실측 전까지 수동 8 h를 그대로 적용했습니다. ARS 효과는 Capa.에 더하지 않았습니다.
충진·포장 1~1.5일, 18 h → 9 h는 검토 중인 개선 방향으로, 계산에는 반영하지 않았습니다.""")


# ============================================================ 4. Reflux Column 원리
from pptx.enum.shapes import MSO_CONNECTOR


def arrow(slide, x1, y1, x2, y2, col, w=1.5):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = RGBColor.from_string(col); c.line.width = Pt(w)
    ln = c.line._get_or_add_ln()
    etree.SubElement(ln, qn("a:tailEnd")).set("type", "triangle")
    return c


def dline(slide, x1, y, x2, col="A6A6A6"):
    from pptx.enum.dml import MSO_LINE
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y), Inches(x2), Inches(y))
    c.line.color.rgb = RGBColor.from_string(col); c.line.width = Pt(0.75); c.line.dash_style = MSO_LINE.DASH
    return c


def rbox(slide, x, y, w, h, fc, lc, l1, l2, c1=DARK):
    sh = box(slide, x, y, w, h, fc, line=lc)
    tf = sh.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; r = p.add_run(); r.text = l1; runfmt(r, 8.5, True, c1)
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r = p2.add_run(); r.text = l2; runfmt(r, 7, False, "595959")
    return sh


def dot(slide, x, y, good, d=0.17):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    sh.shadow.inherit = False
    fill(sh, "E4DFEC" if good else "EDEDED")
    sh.line.color.rgb = RGBColor.from_string(PUR if good else "A6A6A6"); sh.line.width = Pt(0.75)
    return sh


sR = new_slide()
header(sR, 4, "To-be | Reflux Column — 원리와 운전 포인트", "끓는 온도 차이로 탑 안에서 증류를 수십 번 반복 → 반도체용 고순도 프리커서")
sec(sR, 1.05, "① 탑 구조와 흐름", "김 ↑ · 액체 ↓", w=4.2)
CX, CW_ = 1.2, 2.15                       # 탑 위치 · 폭
rbox(sR, CX, 1.36, CW_, 0.46, "DEEBF7", BLU, "Condenser (식히는 곳)", "김을 다시 액체로")
rbox(sR, 3.6, 1.36, 0.95, 0.46, "E2EFD9", GRN, "제품", "일부만 꺼냄")
arrow(sR, CX + CW_, 1.59, 3.6, 1.59, GRN, 1.25)
TOP, SH = 2.02, 0.8                       # 탑 상단 · 단 높이 (4단)
col_ = box(sR, CX, TOP, CW_, SH * 4, "FFFFFF", line="A6A6A6")
pat = ["1111111111", "1111110110", "1010101010", "0100100100"]   # 위 → 아래, 원하는 물질 10 · 8 · 5 · 3개
cnt = [10, 8, 5, 3]
for k in range(4):
    y0 = TOP + k * SH
    if k:
        dline(sR, CX, y0, 4.55)
    for j, ch in enumerate(pat[k]):
        dot(sR, CX + 0.42 + (j % 5) * 0.29, y0 + 0.17 + (j // 5) * 0.3, ch == "1")
    tbox(sR, 3.45, y0 + 0.2, 1.1, 0.42, [[(f"{cnt[k]} / 10", 10, True, PUR)], [("원하는 물질", 6.8, False, "595959")]], align="c")
arrow(sR, CX + 0.17, TOP + SH * 4 - 0.05, CX + 0.17, TOP + 0.05, ORG, 2)          # 김 ↑
arrow(sR, CX + CW_ - 0.17, TOP + 0.05, CX + CW_ - 0.17, TOP + SH * 4 - 0.05, BLU, 2)  # 액체 ↓
arrow(sR, CX + 0.45, TOP, CX + 0.45, 1.82, ORG, 1.5)
arrow(sR, CX + CW_ - 0.45, 1.82, CX + CW_ - 0.45, TOP, BLU, 1.5)
RB = TOP + SH * 4 + 0.2
arrow(sR, CX + CW_ - 0.45, TOP + SH * 4, CX + CW_ - 0.45, RB, BLU, 1.5)
arrow(sR, CX + 0.45, RB, CX + 0.45, TOP + SH * 4, ORG, 1.5)
rbox(sR, CX, RB, CW_, 0.46, "FDE9E7", "EB002C", "Reboiler (냄비)", "데워서 김을 만듦")
rbox(sR, 3.6, RB, 0.95, 0.46, "F2F2F2", "A6A6A6", "찌꺼기", "바닥에 모임")
arrow(sR, CX + CW_, RB + 0.23, 3.6, RB + 0.23, GRAYT, 1.25)
tbox(sR, 0.31, TOP + 0.15, 0.85, 0.5, [[("꼭대기", 8, True, DARK)], [("거의 다", 6.8, False, "595959")], [("원하는 물질", 6.8, False, "595959")]], align="c")
tbox(sR, 0.31, TOP + SH * 1.75, 0.85, 0.5, [[("↑", 12, True, GRAYT)], [("올라갈수록", 6.8, True, "595959")], [("깨끗", 6.8, True, "595959")]], align="c")
tbox(sR, 0.31, TOP + SH * 3.15, 0.85, 0.5, [[("바닥 쪽", 8, True, DARK)], [("불순물이", 6.8, False, "595959")], [("많이 섞임", 6.8, False, "595959")]], align="c")
tbox(sR, 3.45, TOP - 0.02, 1.1, 0.16, [("점선 = 검문소", 6.3, True, REDC)], align="c")
LG = RB + 0.62
dot(sR, 0.4, LG + 0.02, True, 0.14); tbox(sR, 0.6, LG, 1.6, 0.18, [("원하는 물질 (잘 끓음)", 7, False, DARK)])
dot(sR, 2.45, LG + 0.02, False, 0.14); tbox(sR, 2.65, LG, 1.9, 0.18, [("불순물 (잘 안 끓음)", 7, False, DARK)])
arrow(sR, 0.38, LG + 0.33, 0.58, LG + 0.33, ORG, 2); tbox(sR, 0.6, LG + 0.24, 1.6, 0.18, [("김 (위로)", 7, False, DARK)])
arrow(sR, 2.43, LG + 0.33, 2.63, LG + 0.33, BLU, 2); tbox(sR, 2.65, LG + 0.24, 1.9, 0.18, [("액체 (아래로 · Reflux)", 7, False, DARK)])

RX, RW = 4.78, 4.9
sec(sR, 1.05, "② 원리 — 탑 안에서 일어나는 일", x=RX, w=RW)
rows = [["구분", "내용"],
        ["증류", "끓는 온도가 낮은 물질이 먼저 김(증기)이 되어 올라가고, 식히면 다시 액체가 됨"],
        ["한 번으로는 부족", "끓는 온도가 비슷한 불순물은 김에 섞여 같이 올라옴 (체로 한 번 거르는 것과 같음)"],
        ["Reflux Column", "탑 하나 안에서 거르기를 수십 번 반복 → 꼭대기는 고순도, 바닥에는 찌꺼기"],
        ["Reboiler (아래)", "냄비처럼 액체를 데워 김을 만듦"],
        ["Condenser (위)", "올라온 김을 차갑게 식혀 액체로 되돌림"],
        ["Reflux (되돌려 보내기)", "식힌 액체를 전부 꺼내지 않고 일부만 제품으로, 나머지는 탑 안으로 다시 흘려보냄"],
        ["Packing (가운데)", "내려오는 액체와 올라오는 김이 계속 부딪힘 → 잘 끓는 물질은 김 쪽, 안 끓는 물질은 액체 쪽으로"]]
mk_table(sR, RX, 1.33, [1.4, 3.5], rows, fs=7.8, left_cols=(0, 1), wrap=True, bolds={(3, -1)}, fills={(3, -1): "FFF2E6"},
         rh_list=[0.27] + [0.4] * 7)
sec(sR, 4.55, "③ 운전 포인트", x=RX, w=RW)
rows = [["항목", "내용", "생산 영향"],
        ["Reflux 양", "많이 되돌릴수록 여러 번 걸러져 더 깨끗", "꺼내는 양 감소 → 시간·전력 증가 · 순도와 속도의 적정점"],
        ["Packing", "Random(고리 조각): 저렴, 벽 쏠림 → 재분배 장치 / Structured(주름 금속판): 고가, 더 고르게", "분리 성능 · 투자비"],
        ["진공 운전", "프리커서는 열에 약함 → 압력을 낮춰 낮은 온도에서 끓임", "열 분해 방지"],
        ["나눠 담기", "처음(가벼운 불순물) 버림 · 중간(깨끗한 부분)만 제품 · 바닥 찌꺼기(금속 성분) 버림", "회수량 · 수율"]]
mk_table(sR, RX, 4.83, [0.9, 2.55, 1.45], rows, fs=7.6, left_cols=(0, 1, 2), wrap=True, rh_list=[0.27, 0.4, 0.52, 0.4, 0.45])
notes(sR, """Reflux Column은 처음 만들어진 액체에 섞인 원하는 물질과 불순물을 끓는 온도 차이로 골라내는 키 큰 탑입니다. 잘 끓는 물질은 김이 되어 올라가고, 김을 식히면 다시 액체가 됩니다(증류).
한 번 증류로는 끓는 온도가 비슷한 불순물이 같이 올라오므로, 탑 하나 안에서 이 거르기를 수십 번 반복합니다. 아래 Reboiler가 액체를 데워 김을 만들고, 위 Condenser가 김을 식혀 액체로 되돌리며, 그 액체의 일부만 제품으로 꺼내고 나머지는 탑 안으로 다시 흘려보냅니다(Reflux).
탑 가운데 Packing에서 내려오는 액체와 올라오는 김이 계속 만나 단마다 원하는 물질의 비율이 높아집니다(그림: 아래부터 10개 중 3 · 5 · 8 · 10개). 실제 탑에는 이런 단이 수십 개 있어 꼭대기에는 아주 깨끗한 물질만 남고, 불순물은 바닥 Reboiler에 찌꺼기로 모입니다.
Reflux를 많이 할수록 더 깨끗해지지만 꺼내는 양이 줄어 시간이 오래 걸리고 전력도 많이 듭니다. 얼마나 깨끗하게, 얼마나 빨리 만들지의 적정점을 찾는 것이 중요하며, 생산팀 확인 항목인 '환류 운전 시 Batch 시간'과 연결됩니다.
Packing은 고리 모양 조각을 채우는 Random packing(저렴, 벽 쪽 쏠림을 막는 재분배 장치 필요)과 주름진 금속판을 쌓는 Structured packing(고가, 더 고르게 접촉)이 있습니다.
프리커서는 열에 약해 탑 안의 압력을 낮춰(진공) 낮은 온도에서 끓이고, 처음 나오는 가벼운 불순물(초류)과 바닥 찌꺼기(금속 성분)는 버리고 중간의 깨끗한 부분만 제품으로 담습니다.""")

# ============================================================ 11. 생산팀 협의 사항
cw = wb["10_확인사항"]
sM = new_slide()
header(sM, 12, "생산팀 협의 사항 — Batch 기준 · Case", "198 kg/Batch · 1대 월 12 / 2대 월 20 Batch · Case I·II·III")
sec(sM, 1.05, "① 협의 내용과 자료 반영")
rows = [["구분", "협의 내용", "자료 반영", "확인 필요"]]
for rr_ in range(19, 23):
    rows.append([cw.cell(row=rr_, column=c).value for c in range(2, 6)])
mk_table(sM, 0.31, 1.33, [1.3, 3.3, 2.4, 2.37], rows, fs=8, left_cols=(0, 1, 2, 3), wrap=True, rh_list=[0.3] + [0.62] * 4)
sec(sM, 4.3, "② 협의 기준값")
rows = [["Batch size", "Yield", "kg/Batch", "Batch time", "원단위", "1대 월 Batch", "2대 월 Batch", "대정비 월"],
        [f"{V('B_SIZE'):.0f} kg", f"{V('B_YIELD') * 100:.0f}%", f"{KGB:.0f} kg", f"{V('B_TIME'):.0f} h", f"{V('UNIT_RM'):.3f}",
         f"{V('MB_ONE'):.0f} ({f0(V('MB_ONE') * KGB)} kg)", f"{V('MB_TWO'):.0f} ({f0(V('MB_TWO') * KGB)} kg)", f"{V('MB_MAINT'):.0f} ({f0(V('MB_MAINT') * KGB)} kg)"]]
mk_table(sM, 0.31, 4.58, [1.0, 0.8, 1.0, 1.0, 0.85, 1.6, 1.6, 1.52], rows, rh=0.4, fs=8.5, left_cols=(), bolds={(1, -1)})
notes(sM, """오늘 생산팀과 협의한 내용입니다. Batch size 200 kg에 Yield 99%를 적용해 Batch당 198 kg으로 생산량을 계산하고, Batch time 55 h(IQC를 뺀 준비·투입~FQC와 같은 값으로 보임)와 원단위 1.010을 기준으로 삼았습니다.
재고일(개월)은 협의 자료 값을 그대로 썼으며, 재고 ÷ 판매로는 같은 값이 나오지 않아 산식을 확인해야 합니다. 2026년 판매 합계와 고객별 물량의 차이(10·12월 140 kg), 26년 SKH 연간 합계 표기 차이(9,960 / 17,210)도 확인 항목입니다.""")

# ============================================================ 12. 생산팀 확인 사항
sC = new_slide()
header(sC, 13, "생산팀 확인 사항", "생산팀 대화 내용과 추가 확인 항목")
sec(sC, 1.05, "① 생산·설비·품질·충진 관련 확인 내용")
rows = [["구분", "생산팀 설명", "확인 필요"]]
for rr_ in range(6, 19):
    rows.append([cw.cell(row=rr_, column=c).value for c in (2, 3, 5)])
mk_table(sC, 0.31, 1.33, [1.35, 4.75, 3.27], rows, fs=7.6, left_cols=(0, 1, 2), wrap=True, rh_list=[0.28] + [0.395] * 13)
notes(sC, """생산팀과 나눈 대화 중 생산·품질·설비·충진 관련 내용만 정리했습니다. 구두로 언급된 회수량(160~170 kg, 약 220 kg, 195~200 kg)은 운전 조건과 대상이 구분되지 않아 확정 생산량으로 쓰지 않았고, 생산량은 협의 기준 198 kg/Batch를 적용했습니다.
수율은 신규 Crude만이 아니라 재투입·Mix를 포함한 총 투입량 기준으로 따로 계산해야 합니다(Excel 10 시트 입력란). '135 · 120'은 단위가 확인되지 않아 생산량으로 입력하지 않았고, '65%에서 5% 상승'도 정의를 확인할 항목입니다.
이지켐은 색도 때문에 추가 투입을 제한하는 경우가 있어, 고객별 품질·색도 규격과 합격률을 별도로 확인합니다.""")

# ============================================================ 순서 정리 · 원본 3~5장 제거
order = [s1, sP, sW, sR, s2, sB, sV] + case_slides + [sL, sM, sC]
lst = prs.slides._sldIdLst
ids = {prs.part.related_part(x.rId): x for x in lst}
for old in (s3, s4, s5):
    x = ids[old.part]; prs.part.drop_rel(x.rId); lst.remove(x)
for s in order:
    x = ids[s.part]; lst.remove(x); lst.append(x)
prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides")
