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

from pptx_helpers import (E, FONT, RGBColor, box, clone, copy_fill, delete, etree, notes, pos, runfmt, set_lines, shapes, tbox,
                          widen_flow)

BASE, XLSX, MAPF, OUT = sys.argv[1:5]
MAP = json.load(open(MAPF))
wb = openpyxl.load_workbook(XLSX, data_only=True)
REF = MAP["REF"]


def V(k):
    s, c = REF[k]
    return wb[s][c].value


VS = {"A": "05_월별_Capa기준", "B": "06_월별_생산계획"}


def mrow(ver, yr, key):
    ws = wb[VS[ver]]; r = MAP["MROW"][ver][str(yr)][key]
    return [ws.cell(row=r, column=2 + m).value for m in range(1, 13)], ws.cell(row=r, column=15).value


def lrow(yr, key):
    ws = wb["08_후공정부하"]; r = MAP["LOAD"][str(yr)][key]
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


def sec(slide, y, main, sub="", x=0.31):
    bar, txt = (copy.deepcopy(e) for e in SEC)
    for el in (bar, txt):
        slide.shapes._spTree.append(el)
        el.xpath("./*[1]/p:cNvPr")[0].set("id", str(max(s_.shape_id for s_ in slide.shapes) + 1))
    b, t = slide.shapes[-2], slide.shapes[-1]
    pos(b, x=x, y=y + 0.03); pos(t, x=x + 0.12, y=y, w=9.2 - x)
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
def gantt(slide, top, units, intv, share200, lanes, pitch=0.29):
    """units: 정제기별 첫 투입 시점(h) 목록. 완료(FQC 종료)가 0~720 h 안이면 이 달 생산으로 계산."""
    X0, X1, H = 1.32, 9.69, 720.0
    kk = (X1 - X0) / H
    t_iqc, t_prep, t_ref, t57 = V("T_IQC"), V("T_PREP"), V("T_REF"), V("T_57")
    f5, f2 = V("FILL_5G") * V("BOT_B"), V("FILL_EZ")
    for day in range(1, 31):
        tbox(slide, X0 + (day - 1) * 24 * kk, top, 24 * kk, 0.15, [(str(day), 6.8, day in (1, 30), "404040")], align="c")
    tbox(slide, 0.31, top, 1.0, 0.15, [("일 (30일)", 6.8, False, GRAYT)])
    lt = top + 0.17
    nl = len(lanes)
    for i, lab in enumerate(lanes):
        tbox(slide, 0.31, lt + i * pitch + (pitch - 0.17) / 2, 1.0, 0.17, [(lab, 7.8, True, DARK)])
        box(slide, X0, lt + (i + 1) * pitch - 0.02, X1 - X0, 0.004, "E7E6E6")
    for day in range(31):
        box(slide, X0 + day * 24 * kk, lt, 0.004, nl * pitch - 0.02, "E7E6E6")
    ly = {lab: lt + i * pitch + (pitch - 0.26) / 2 for i, lab in enumerate(lanes)}

    def bar(lane, a, b, color, text="", tcol="FFFFFF", **kw):
        a, b = max(a, 0), min(b, H)
        if b <= a:
            return
        box(slide, X0 + a * kk, ly[lane], (b - a) * kk, 0.24, color, text if (b - a) * kk > 0.3 else "", 6.6, tcol, **kw)

    done = []
    for u, ph in enumerate(units):
        lane = lanes[u]
        for k in range(-2, 16):
            st = ph + k * intv
            if st > H:
                break
            lab = "전월" if st + t_iqc + t_prep < 0 else (f"B{k + 1}" if len(units) == 1 else f"{u + 1}-{k + 1}")
            bar(lane, st + t_iqc, st + t_iqc + t_prep + t_ref, ORG, lab)
            bar(lane, st + t_iqc + t_prep + t_ref, st + intv + t_iqc, None, line="BF9000", dash=True)
            bar(lanes[len(units)], st + t_iqc + t_prep + t_ref, st + t57, BLU)
            done.append(st + t57)
    done = sorted(c for c in done if c > -f5)
    cnt = [c for c in done if 0 <= c < H]
    n200 = round(len(cnt) * share200)
    idx = 0; cum = 0
    for c in done:
        if 0 <= c < H:
            is200 = int((idx + 1) * n200 / len(cnt)) > int(idx * n200 / len(cnt))
            idx += 1; cum += V("KG_B")
            tbox(slide, X0 + c * kk - 0.25, ly[lanes[-1]] + 0.04, 0.5, 0.16, [(f"{cum:,.0f}", 6.6, True, BLU)], align="c")
        else:
            is200 = False
        if is200:
            bar(lanes[-2], c, c + f2, PUR, "200 L")
        else:
            bar(lanes[-3], c, c + f5, GRN, "5 Gal")
    yl = lt + nl * pitch + 0.04
    xl = 1.32
    for colr, lab, dash in ((ORG, f"준비·정제 {t_prep + t_ref:.0f} h", False), (None, "해제·전환 (확인)", True),
                            (BLU, "PQC~FQC 8 h", False), (GRN, f"5 Gal 충진 {f5:.0f} h", False), (PUR, f"200 L 충진 {f2:.0f} h", False)):
        box(slide, xl, yl + 0.03, 0.2, 0.12, colr, line="BF9000" if dash else None, dash=dash)
        tbox(slide, xl + 0.25, yl, 1.4, 0.17, [(lab, 7.2, False, "404040")])
        xl += 1.68
    return len(cnt), yl + 0.2


CHAN = {}
for yr, m in ((2027, 1), (2027, 7)):
    hx, cx, ez, hs = (mrow("A", yr, k)[0][m - 1] for k in ("hx", "cx", "ez", "hs"))
    CHAN[(yr, m)] = (ez + hs) / (hx + cx + ez + hs)

CM, CM105, MK, KGB = V("CAPA_M"), V("CAPA_M105"), V("MAINT_KG"), V("KG_B")
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
sec(s1, 3.80, "① 공정시간 단순 합계 (h)", "IQC~FQC 57 h + 충진 + OQC")
x0, k = 2.10, (8.85 - 2.10) / 77
paths = (("5 Gal 9병 (하이닉스·CXMT)", 18, "충진 18 h"), ("이지켐 200 L 1용기", 8, "충진 8 h"), ("한솔 200 L 1용기", 8, "충진 8 h"))
segs = ((V("T_IQC"), "A6A6A6", "IQC 2"), (V("T_PREP"), "F4B183", "2"), (TREF, ORG, f"순수 정제 {TREF:.0f} h"), (V("T_PQC"), "9DC3E6", "PQC 6"),
        (V("T_FQC"), BLU, "FQC 2"))
for i, (lab, fh, flab) in enumerate(paths):
    yy = 4.12 + i * 0.38
    tbox(s1, 0.31, yy + 0.06, 1.8, 0.18, [(lab, 7.8, True, DARK)])
    xx = x0
    for hh, colr, t in segs + ((fh, GRN, flab), (V("OQC_200") if i else V("OQC_5G"), "7F7F7F", "OQC 2")):
        box(s1, xx, yy, hh * k, 0.3, colr, t if hh * k > 0.45 else "", 7, "1F1F1F" if colr in ("A6A6A6", "F4B183", "9DC3E6") else "FFFFFF")
        xx += hh * k
    tbox(s1, xx + 0.06, yy + 0.04, 0.7, 0.2, [(f"{V(f'lt{i}'):.0f} h", 9.5, True, REDC)])
mx = x0 + T57 * k
box(s1, mx, 4.06, 0.014, 1.16, REDC)
tbox(s1, mx - 0.6, 5.24, 1.2, 0.17, [(f"▲ IQC~FQC {T57:.0f} h", 7.5, True, REDC)], align="c")
rows = [["경로", "IQC~FQC", "충진", "OQC·출하", "합계", "충진량 · 기준"]]
for i, (lab, fh, _) in enumerate(paths):
    rows.append([lab, f"{T57:.0f} h", f"{fh} h", "2 h", f"{V(f'lt{i}'):.0f} h",
                 ["180 kg (9병) · 2 h/병", "140 kg · 8 h/용기", "150 kg · 8 h/용기"][i]])
mk_table(s1, 0.31, 5.55, [2.6, 1.15, 1.15, 1.15, 1.15, 2.17], rows, rh=0.3, fs=8.5,
         bolds={(1, 4), (2, 4), (3, 4)}, colors={(i, 4): REDC for i in (1, 2, 3)})
notes(s1, f"""As-is는 정제기 1대로 하이닉스·CXMT·이지켐(한솔 '27~)을 함께 생산합니다.
57 h는 IQC부터 FQC까지의 현재 보고 기준이고, 순수 정제 {TREF:.0f} h는 57 − IQC 2 − 준비·투입 2 − PQC~이송 6 − FQC 2로 계산한 값입니다(실측 아님). 대화 중 57 − 2 − 6 = 49 h 계산, 과거 약 72 h · 최근 약 53 h 언급이 있어 원자료의 측정 시작·종료점은 확인이 필요합니다.
PQC와 제품 이송은 합쳐서 6 h입니다. 57 h에는 충진·OQC가 포함되지 않습니다.
5 Gal 9병은 2 h × 9 = 18 h 작업이며 현장에서는 근무시간 기준 약 2일로 설명합니다. 200 L은 이지켐·한솔 모두 수동 8 h/용기로 반영했습니다.
경로별 합계 77 h · 67 h는 확인된 시간 항목을 더한 값으로, 달력 기준 납기나 다음 Batch 투입 간격과는 다릅니다. 충진이 진행되는 동안 다음 Batch 정제가 가능합니다.
190 kg과 5 Gal 9병 180 kg의 차이 10 kg은 잔량·추가 충진·다른 고객 배분 여부를 확인할 항목입니다.""")

# ============================================================ 2. As-is 한 달 운전
sA = new_slide()
ai, ab, akg = V("asis_int"), V("asis_b"), V("asis_kg")
header(sA, 2, "As-is | 정제기 1대 한 달 운전", f"정제기 1대 한 달 {ab:.0f} Batch = {f0(akg)} kg (과거 최대) · 연간 26.2톤 = 월 {f0(V('CAPA_NOW') * 1000 / 12)} kg")
sec(sA, 1.05, "① 30일 운전 예시 — 과거 최대 월 11 Batch", f"투입 간격 720 ÷ 11 = {ai:.1f} h")
n_, yb = gantt(sA, 1.36, [0.0], ai, CHAN[(2027, 1)], ["정제기", "PQC~FQC", "5 Gal 충진", "200 L 충진", "누적 kg"], pitch=0.38)
cm = wb["03_한달운전"]
cr = lambda r_, c_: cm.cell(row=r_, column=c_).value
sec(sA, 3.86, "② 월 생산량 비교 — 정제기 1대")
rows = [["기준", "Batch/월", "kg/월", "연간 (t)", "근거"]]
for rr_, lab, basis in ((42, "이론 (순수 정제만)", "720 h ÷ 45 h"), (43, "과거 최대", "11 Batch × 190 kg"), (44, "현재 Capa. 26.2톤", "26,200 kg ÷ 12")):
    rows.append([lab, f"{cr(rr_, 4):.1f}" if rr_ == 44 else f"{cr(rr_, 4):.0f}", f0(cr(rr_, 5)), f"{cr(rr_, 6):.1f}", basis])
rows.append(["2026 출하 (제시분)", f"{V('sh26') / 12 / KGB:.1f}", f0(V("sh26") / 12), f"{V('sh26') / 1000:.1f}", "CXMT·한솔 미제시"])
mk_table(sA, 0.31, 4.14, [2.6, 1.1, 1.2, 1.2, 3.27], rows, rh=0.32, fs=8.5, bolds={(2, -1), (3, -1)}, left_cols=(0, 4),
         fills={(2, -1): "FFF2E6"})
sec(sA, 5.86, "③ 과거 최대 11 Batch의 720 h 구성")
k7 = (9.69 - 1.32) / 720
oth = 720 - 11 * (TREF + V("T_PREP"))
segs = ((11 * TREF, ORG, f"순수 정제 11 × {TREF:.0f} = {11 * TREF:.0f} h"), (11 * V("T_PREP"), "F4B183", f"{11 * V('T_PREP'):.0f}"),
        (oth, "D9D9D9", f"해제·전환·대기·보수 등 {oth:.0f} h"))
xx = 1.32
tbox(sA, 0.31, 6.22, 1.0, 0.2, [("720 h", 8.5, True, DARK)])
for hh, colr, t in segs:
    box(sA, xx, 6.16, hh * k7, 0.36, colr, t, 8, "1F1F1F" if colr != ORG else "FFFFFF")
    xx += hh * k7
notes(sA, f"""정제기 1대의 과거 최대인 월 11 Batch를 30일(720 h)에 놓으면 투입 간격은 {ai:.1f} h이고, 11 × 190 = {f0(akg)} kg입니다.
막대는 준비·투입 2 h + 순수 정제 {TREF:.0f} h를 정제기 운전으로, 이후 다음 투입까지를 해제·전환 구간(구성 확인)으로 표시했습니다. 정제가 끝난 Batch는 PQC~이송·FQC 8 h 뒤 충진하며, 그동안 다음 Batch가 정제됩니다.
순수 정제만 놓고 보면 720 ÷ 45 = 16 Batch(3,040 kg)이지만 점유·전환·보수 등을 뺀 이론값입니다. 11 Batch 기준 495 h를 뺀 225 h를 전부 손실이나 충진시간으로 보지 않습니다.
현재 연간 Capa. 26.2톤은 190 kg 기준 137.9 Batch 상당(월 11.5 Batch, 2,183 kg)입니다. 운전시간·가동률을 조정해서 맞춘 값이 아닙니다.
2026 출하 제시분은 19,220 kg(하이닉스 16,980 + 이지켐 2,240)이며 CXMT·한솔 물량은 아직 제시되지 않았습니다.""")

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
header(s2, 3, "To-be | 정제기 2대 운영과 투자 일정", "정제기 2대의 투입 시점을 엇갈리게 운영하여 생산을 병행하고, 후공정 부하를 분산하는 계획")
sec(s2, 3.76, "① 투자 · 운영 일정", "매년 4월 대정비 · '27.5 정제기 2대 생산 시작")
yrs = ["구분", "2027년"] + [""] * 11 + ["2028년"] + [""] * 5
mons = [""] + [str(m) for m in range(1, 13)] + [str(m) for m in range(1, 7)]
r_cur = ["정제기 1대 (확정)", "2,178 · 2,178 · 3,168", "", ""] + [""] * 15
r_mt = ["대정비 (매년 4월)"] + [""] * 3 + [f"{f0(MK)}"] + [""] * 11 + [f"{f0(MK)}"] + ["", ""]
r_rf = ["정제기 2대 생산", "", "", "", "", f"월 {f1(CM)} kg (47,200 ÷ 11) · 기존 생산계획 3,564 kg"] + [""] * 7 + [f"월 {f1(CM105)} kg", "", "", "", f"{f1(CM105)}", ""]
r_ars = ["200 L 충진", "수동 8 h/용기"] + [""] * 5 + ["ARS '27.7~ (이지켐 · 한솔) · 5 Gal은 글로브 박스 유지"] + [""] * 11
r_105 = ["105℃ (2028)"] + [""] * 12 + ["103 → 105℃ · +3톤 (50.2톤)"] + [""] * 5
rows = [yrs, mons, r_cur, r_mt, r_rf, r_ars, r_105]
fl = {(0, -1): "595959", (1, -1): "F2F2F2"}
fl.update({(2, j): "FDE9E7" for j in (1,)}); fl.update({(3, 4): "D9D9D9", (3, 16): "D9D9D9"})
fl.update({(4, 5): "FDE9E7", (4, 13): "FDE9E7", (4, 17): "FDE9E7", (5, 7): "E4DFEC", (6, 13): "E4DFEC"})
cl = {(1, j): DARK for j in range(19)}
cl.update({(2, 1): REDC, (3, 4): REDC, (3, 16): REDC, (4, 5): REDC, (4, 13): REDC, (4, 17): REDC, (5, 1): GRAYT, (5, 7): PUR, (6, 13): PUR})
mk_table(s2, 0.31, 4.02, [1.6] + [0.43] * 18, rows, rh=0.235, fs=7.3,
         merges=[(0, 1, 12), (0, 13, 18), (2, 1, 3), (4, 5, 12), (4, 13, 15), (4, 17, 18), (5, 1, 6), (5, 7, 18), (6, 13, 18)],
         fills=fl, colors=cl, bolds={(2, 1), (3, 4), (3, 16), (4, 5), (4, 13), (4, 17), (1, -1)}, hdr_rows=1)
sec(s2, 5.78, "② 월 생산 기준 — 4월 대정비 1,188 kg · 나머지 11개월에 연간 Capa. ÷ 11")
rows = [["구분", "연간 Capa.", "4월 (대정비)", "나머지 11개월", "산식"],
        ["현재 정제기 1대", f"{V('CAPA_NOW'):.1f}톤", "-", f"{f0(V('CAPA_NOW') * 1000 / 12)} kg", "26,200 ÷ 12"],
        ["정제기 2대 ('27.5~)", f"{V('CAPA_RF'):.1f}톤", f"{f0(MK)} kg", f"{f1(CM)} kg", "47,200 ÷ 11"],
        ["2028 105℃ 적용", f"{V('CAPA_105'):.1f}톤", f"{f0(MK)} kg", f"{f1(CM105)} kg", "50,200 ÷ 11"]]
mk_table(s2, 0.31, 6.04, [2.6, 1.3, 1.5, 1.7, 2.27], rows, rh=0.205, fs=8, bolds={(2, -1), (3, -1)},
         colors={(2, 3): REDC, (3, 3): REDC})
notes(s2, f"""리플럭스 이후에는 정제기 2대를 활용합니다. 정제기 1을 먼저 투입·운전하고, 초류 진행에 맞춰 정제기 2를 투입해 두 설비의 운전 구간을 겹칩니다(현장 설명: 다음 날 투입·초류 중간·초류 종료 무렵 — 투입 간격은 미확정).
정제기별 Batch 시간이 절반으로 줄어드는 것이 아니라 두 설비의 생산을 겹쳐 전체 생산량을 늘리는 방식이며, 제품 회수 시점이 엇갈려 검사·충진 부하가 한꺼번에 몰리지 않습니다. 57 ÷ 2 = 28.5 h는 이상적인 경우의 도착 간격 설명일 뿐 정제기별 Batch 시간이 아닙니다.
일정: '27.1~3 정제기 1대 확정 생산(2,178 · 2,178 · 3,168 kg) → '27.4 대정비({f0(MK)} kg) → '27.5부터 정제기 2대. 대정비는 매년 4월로 반영했습니다. 200 L ARS는 '27.7부터(이지켐·한솔), 5 Gal은 글로브 박스를 유지합니다.
Capa.: 26.2 + 21 = 47.2톤/년. 월 기준은 47,200 ÷ 11 = {f1(CM)} kg이고 대정비 월(4월)은 {f0(MK)} kg을 따로 둡니다. 2028년은 105℃ 적용(+3톤)으로 50.2톤, 50,200 ÷ 11 = 월 {f1(CM105)} kg — 열 안정성·Dimer·Unknown impurity·수율 검증과 승인이 전제입니다.
47.2톤이 현재의 2배(52.4톤)가 아닌 것은 추가 투입용 Mix·Premix 준비 등에 설비 시간이 필요하기 때문이며, 각 요인의 손실량은 아직 정량화되지 않았습니다.
리플럭스는 하이닉스 요구 품질 대응과 생산능력 확대를 함께 보는 프로젝트입니다. 추가 정제기가 기존 설비 개조인지 신규 설치인지, 설비 번호(2070·2060 등)는 도면과 대조해 확인합니다.""")

# ============================================================ 4. To-be 한 달 운전
sB = new_slide()
ti, toff, tb, tbm = V("tobe_int"), V("tobe_off"), V("tobe_b"), V("tobe_bm")
header(sB, 4, "To-be | 정제기 2대 시차 운전 한 달", f"정제기 2대 · 30일 {tb:.0f} Batch = {f0(tb * KGB)} kg · 월평균 {tbm:.1f} Batch = {f0(CM)} kg (47.2톤)")
sec(sB, 1.05, "① 30일 운전 예시 — 정제기 2대 투입 시점 엇갈림", f"설비별 투입 간격 {ti:.1f} h · 정제기 2 시차 {toff:.0f} h (가정)")
gantt(sB, 1.36, [0.0, toff], ti, CHAN[(2027, 7)], ["정제기 1", "정제기 2", "PQC~FQC", "5 Gal 충진", "200 L 충진", "누적 kg"], pitch=0.36)
sec(sB, 4.12, "② 월 생산량 비교 (대정비 외 달)")
rows = [["기준", "정제기", "Batch/월", "kg/월", "연간 (t)", "설비별 간격 (h)"]]
for rr_, lab in ((44, "현재 Capa. 26.2톤"), (43, "과거 최대"), (45, "기존 생산계획 ('27.5~)"), (46, "Capa. 47.2톤 기준"), (47, "2028 105℃ 50.2톤 기준")):
    yv = cr(rr_, 6)
    rows.append([lab, f"{cr(rr_, 3):.0f}대", f"{cr(rr_, 4):.1f}", f0(cr(rr_, 5)), f"{yv:.1f}" if isinstance(yv, (int, float)) else "-", f"{cr(rr_, 7):.1f}"])
pl27 = mrow("B", 2027, "plan")[1]
rows[3][4] = f"{pl27 / 1000:.1f} ('27)"
mk_table(sB, 0.31, 4.40, [2.9, 0.9, 1.2, 1.3, 1.3, 1.77], rows, rh=0.37, fs=8.5, bolds={(4, -1), (5, -1)},
         fills={(4, -1): "FFF2E6", (5, -1): "FFF2E6"})
notes(sB, f"""정제기 2대 운전 예시입니다. 설비별 투입 간격 {ti:.1f} h는 월 Capa. {f1(CM)} kg ÷ 190 = {tbm:.1f} Batch를 평균 월 730 h에 2대로 나눈 역산값이고, 정제기 2 시차 {toff:.0f} h는 설명용 가정입니다. 실제 투입 간격은 초류 진행과 충진 처리능력을 보고 정합니다.
30일(720 h) 안에 FQC가 끝나는 Batch는 {tb:.0f}개({f0(tb * KGB)} kg)이고, 월평균 730 h 기준으로는 {tbm:.1f} Batch = {f0(CM)} kg입니다.
두 정제기의 완료 시점이 엇갈려 5 Gal 충진(18 h)과 200 L 충진(8 h)이 겹치지 않게 배치됩니다. 충진은 출하 비중(7월 기준 200 L 약 {CHAN[(2027, 7)] * 100:.0f}%)에 맞춰 일부 Batch를 200 L로 표시했습니다.
기존 생산계획의 5~12월 3,564 kg은 190 kg 기준 월 18.8 Batch, 설비별 간격 약 {cr(45, 7):.0f} h에 해당합니다. 2028년 105℃ 기준 {f1(CM105)} kg은 월 {CM105 / KGB:.1f} Batch입니다.
Batch 수가 늘어나면 글로브 박스 충진·검사·Tank·포장이 먼저 병목이 될 수 있어, 정제능력과 최종 출하능력은 구분해서 봅니다.""")

# ============================================================ 5. 버전 비교
sV = new_slide()
vw = wb["07_버전비교"]
va = [vw.cell(row=6, column=c).value for c in range(2, 13)]
vb = [vw.cell(row=7, column=c).value for c in range(2, 13)]
INV0 = V("INV_0")
header(sV, 5, "2027~2028 생산 계획 버전 비교",
       f"'28 말 재고 — ① Capa. 기준 {sg(va[8])} kg · ② 기존 생산계획 {sg(vb[8])} kg")
sec(sV, 1.05, "① 연간 생산 · 출하 · 재고 (kg)", f"재고 = 전월 재고 + 생산 − 출하 · 2027.1.1 기초재고 {f0(INV0)} kg")
rows = [["버전", "2027 생산", "2027 출하", "차이", "2028 생산", "2028 출하", "차이", "'27 말 재고", "'28 말 재고", "최저 재고 (시점)"]]
for v_, lab in ((va, "① Capa. 기준 (47.2 / 50.2톤)"), (vb, "② 기존 생산계획")):
    rows.append([lab, f0(v_[1]), f0(v_[2]), sg(v_[3]), f0(v_[4]), f0(v_[5]), sg(v_[6]), sg(v_[7]), sg(v_[8]), f"{sg(v_[9])} ('{v_[10]:%y.%-m})"])
cl = {(i, j): "385723" for i in (1, 2) for j in (3, 6, 7, 8, 9) if rows[i][j].startswith("+")}
mk_table(sV, 0.31, 1.33, [2.2, 0.78, 0.78, 0.72, 0.78, 0.78, 0.72, 0.8, 0.8, 1.01], rows, rh=0.34, fs=8,
         bolds={(i, j) for i in (1, 2) for j in (3, 6, 7, 8, 9)}, colors=cl)
ser = [[vw.cell(row=12 + i, column=3 + k).value for k in range(24)] for i in range(5)]
labs = ("출하", "① 생산", "① 재고", "② 생산", "② 재고")
for bi, (yr, y0) in enumerate(((2027, 2.5), (2028, 4.66))):
    sec(sV, y0, f"{'②③'[bi]} {yr} 월별 생산 · 재고 (kg)")
    rows = [["구분"] + [f"{m}월" for m in range(1, 13)] + ["연간"]]
    for i, lab in enumerate(labs):
        vals = ser[i][bi * 12:(bi + 1) * 12]
        if "재고" in lab:
            rows.append([lab] + [sg(v) for v in vals] + [sg(vals[-1])])
        else:
            rows.append([lab] + [f0(v) if abs(v - round(v)) < 1e-6 else f1(v) for v in vals] + [f0(sum(vals))])
    cl = {(i, j): "385723" for i in (3, 5) for j in range(1, 14) if rows[i][j].startswith("+")}
    mk_table(sV, 0.31, y0 + 0.28, [1.1] + [0.62] * 12 + [0.83], rows, rh=0.29, fs=7.6, colors=cl,
             bolds={(3, -1), (5, -1)}, fills={(1, -1): "F2F2F2", (3, -1): "FFF2E6", (5, -1): "EAF1FB"})
notes(sV, f"""두 버전을 같은 출하계획(2027 {va[2] / 1000:.2f}톤, 2028 {va[5] / 1000:.2f}톤)과 비교했습니다. 2027년 1~4월은 확정 생산(2,178 · 2,178 · 3,168 · 1,188 kg)으로 두 버전이 같고, 재고는 2027.1.1 기초재고 {f0(INV0)} kg에서 시작해 전월 재고 + 생산 − 출하로 이어집니다.
① Capa. 기준: 2027년 5~12월 47,200 ÷ 11 = {f1(CM)} kg/월, 2028년 105℃ 50,200 ÷ 11 = {f1(CM105)} kg/월(4월 대정비 {f0(MK)} kg 별도). 2027년 {sg(va[3])} kg, 2028년 {sg(va[6])} kg. 4월 대정비 달에 재고가 {sg(va[9])} kg까지 내려가고, '27 말 {sg(va[7])} kg, '28 말 {sg(va[8])} kg입니다.
② 기존 생산계획: 2027년 5~12월 3,564 kg, 2028년 3,564 kg(4월 1,188 · 8~12월 3,168 kg). 2027년 {sg(vb[3])} kg, 2028년 {sg(vb[6])} kg으로 '27 말 {sg(vb[7])} kg, '28 말 {sg(vb[8])} kg까지 부족이 커집니다.""")

# ============================================================ 6~9. 연도별 · 버전별
TITLES = {("A", 2027): ("2027년 ① Capa. 기준 (47.2톤)", f"1~3월 확정 · 4월 대정비 {f0(MK)} · 5~12월 47,200 ÷ 11 = {f1(CM)} kg"),
          ("B", 2027): ("2027년 ② 기존 생산계획", f"1~3월 확정 · 4월 대정비 {f0(MK)} · 5~12월 3,564 kg"),
          ("A", 2028): ("2028년 ① Capa. 기준 (50.2톤)", f"4월 대정비 {f0(MK)} · 나머지 11개월 50,200 ÷ 11 = {f1(CM105)} kg"),
          ("B", 2028): ("2028년 ② 기존 생산계획", f"4월 대정비 {f0(MK)} · 1~7월 3,564 · 8~12월 3,168 kg")}
year_slides = []
for num, (ver, yr) in enumerate((("A", 2027), ("B", 2027), ("A", 2028), ("B", 2028)), start=6):
    s = new_slide(); year_slides.append(s)
    g = {k: mrow(ver, yr, k) for k in ("cond", "plan", "pb", "diff", "inv", "hx", "cx", "ez", "hs", "ship", "nb", "l2m")}
    pl, sh_, df, inv = g["plan"][1], g["ship"][1], g["diff"][1], g["inv"][0]
    lowv = min(inv); lowm = inv.index(lowv) + 1
    ttl, basis = TITLES[(ver, yr)]
    stock = f"연말 재고 {sg(inv[-1])} kg" + (f" ({lowm}월 {sg(lowv)})" if (lowv < 0 and lowm < 12) else "")
    header(s, num, ttl, f"생산 {pl / 1000:.2f}톤 vs 출하 {sh_ / 1000:.2f}톤 → {sg(df)} kg · {stock}")
    sec(s, 1.05, f"① 월별 생산 · 출하 · 재고 (kg)", basis)
    conds = g["cond"][0]
    spans, a = [], 0
    for m in range(1, 13):
        if m == 12 or conds[m] != conds[a]:
            spans.append((a + 1, m, conds[a])); a = m
    crow = ["운전 조건"] + [""] * 13
    for a_, b_, t_ in spans:
        crow[a_] = t_
    rows = [["구분"] + [f"{m}월" for m in range(1, 13)] + ["연간"], crow]
    spec = (("plan", "생산", f1, f0), ("pb", "Batch (÷190)", f1, f1), ("diff", "생산 − 출하", sg, sg),
            ("inv", "재고", sg, sg), ("hx", "SK하이닉스", f0, f0), ("cx", "CXMT", f0, f0),
            ("ez", "이지켐", f0, f0), ("hs", "한솔", f0, f0), ("ship", "출하 합계", f0, f0), ("nb", "필요 Batch", f1, f1),
            ("l2m", "200 L 충진", lambda v: v.replace(" ", ""), str))
    for key, lab, fm, ft in spec:
        vals, tot = g[key]
        if key == "plan":
            cells = [f0(v) for v in vals]
        else:
            cells = [fm(v) for v in vals]
        rows.append([lab] + cells + [ft(tot) if key != "inv" else sg(vals[-1])])
    fills = {(0, -1): "595959", (5, -1): "FFF2E6"}
    for a_, b_, t_ in spans:
        fills[(1, a_)] = {"대정비": "D9D9D9", "정제기 1대": "F2F2F2"}.get(t_, "FDE9E7")
    for m in range(1, 13):
        if conds[m - 1] == "대정비":
            fills[(2, m)] = "EDEDED"
        elif yr == 2027 and m <= 4:
            fills[(2, m)] = "DDEBF7"
    colors = {(1, a_): REDC for a_, b_, t_ in spans if t_ == "대정비"}
    for ri in (4, 5):
        for j in range(1, 14):
            if rows[ri][j].startswith("+"):
                colors[(ri, j)] = "385723"
    mk_table(s, 0.31, 1.32, [1.15] + [0.615] * 12 + [0.84], rows, rh=0.38, fs=8.2,
             merges=[(1, a_, b_) for a_, b_, t_ in spans if b_ > a_], fills=fills, colors=colors,
             bolds={(2, -1), (5, -1), (10, -1), (1, -1)})
    ez_m = g["ez"][0]
    notes(s, (f"{ttl}. 생산 {f0(pl)} kg, 출하 {f0(sh_)} kg, 차이 {sg(df)} kg. " +
              ("2027년 1~4월은 확정 생산(2,178 · 2,178 · 3,168 · 1,188 kg, 4월 대정비)입니다. " if yr == 2027 else "") +
              ("5월부터 월 Capa. 47,200 ÷ 11 = " + f1(CM) + " kg을 적용했습니다. " if (ver, yr) == ("A", 2027) else "") +
              ("105℃ 50.2톤 기준으로 50,200 ÷ 11 = " + f1(CM105) + " kg, 4월 대정비 1,188 kg입니다. 105℃는 품질 검증·승인이 전제입니다. "
               if (ver, yr) == ("A", 2028) else "") +
              ("기존 생산계획 수치(3,564 · 3,168 · 2,178 · 1,188 kg)는 198 kg × 정수 Batch(18 · 16 · 11 · 6)와 일치합니다 — 계획 Batch량 기준 확인 필요. "
               if ver == "B" else "") +
              f"재고는 2027.1.1 기초재고 {f0(V('INV_0'))} kg에서 전월 재고 + 생산 − 출하로 이어지며, {yr}년 최저 {sg(lowv)} kg({lowm}월), 연말 {sg(inv[-1])} kg입니다. " +
              f"출하: 하이닉스 22,720 ÷ 12, CXMT {f0(g['cx'][1])}, 이지켐 {f0(g['ez'][1])}, 한솔 {f0(g['hs'][1])} kg. " +
              ("이지켐은 7월부터 월 840 kg입니다. " if yr == 2027 else "2028년 물량은 2027년 하반기 수준을 유지한 비교용 가정입니다. ") +
              "200 L 충진은 이지켐·한솔 수동 8 h/용기 기준이며, ARS 시간이 실측되면 Excel 01_입력에서 바뀝니다."))

# ============================================================ 10. 후공정 부하
sL = new_slide()
l27 = {k: lrow(2027, k)[0] for k in ("g5bt", "g5h", "gbav", "gbld", "ezc", "hsc", "l2h", "g5oqc", "l2oqc", "qch", "nb")}
l28 = {k: lrow(2028, k)[0] for k in l27}
peak = max(l27["gbld"] + l28["gbld"])
header(sL, 10, "후공정 부하 검토 — 충진 · 검사 · OQC", f"정제능력과 출하능력은 별개 · 5 Gal 충진 최대 월 {max(l27['g5h'] + l28['g5h']):.0f} h = 1교대 {l27['gbav'][0]:.0f} h의 {peak * 100:.0f}%")
sec(sL, 1.05, "① 월 작업량 (출하계획 기준)", "출하 물량으로 계산 — 생산 버전과 무관")
cols = (("'27.1~2", l27, 0), ("'27.3~6", l27, 2), ("'27.7~12", l27, 6), ("'28 (월)", l28, 0))
items = (("5 Gal 병", "g5bt", f0, "하이닉스 + CXMT ÷ 20 kg"), ("5 Gal 충진 h", "g5h", f0, "2 h/병"),
         ("글로브 박스 가용 h", "gbav", f0, "1교대 8 h × 22일"), ("글로브 박스 부하율", "gbld", lambda v: f"{v * 100:.0f}%", "1교대 기준"),
         ("200 L 용기", None, f0, "이지켐 140 · 한솔 150 kg"), ("200 L 충진 h", "l2h", f0, "8 h/용기"),
         ("OQC·출하 h", None, f0, "9병 2 h · 1용기 2 h"), ("검사 h (PQC~FQC)", "qch", f0, "8 h × 필요 Batch"),
         ("필요 Batch", "nb", f1, "출하 ÷ 190"))
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
        ["5 Gal 글로브 박스", "9병 18 h (약 2일)", "월 267 h > 1교대 176 h", "교대 · 인원 · 글로브 박스 수"],
        ["200 L 충진", "수동 8 h/용기", "ARS '27.7~ (시간 비슷할 수 있음)", "ARS 단계별 시간 실측"],
        ["검사 (PQC · FQC)", "8 h/Batch", "Batch 증가 · 회수 시점 겹침", "검사 인력 · 승인 대기"],
        ["Product Tank", "PQC 합격 후 이송", "정제기 2대 제품 보관", "Tank 수 · 용량"],
        ["충진 · 포장", "1 Batch 약 2일", "목표 1~1.5일", "개선 방안 · 일정"]]
mk_table(sL, 0.31, 4.58, [1.9, 2.0, 2.7, 2.77], rows, rh=0.35, fs=8, left_cols=(0, 1, 2, 3))
notes(sL, f"""후공정 작업량은 출하 물량으로 계산했습니다. 5 Gal은 (하이닉스 + CXMT) ÷ 20 kg × 2 h, 200 L은 이지켐 140 kg · 한솔 150 kg 용기 × 8 h(수동 기준)입니다.
2027년 하반기부터 5 Gal 충진이 월 약 267 h로 1교대(8 h × 22일 = 176 h)를 넘습니다. 정제기가 2대가 되어도 충진·검사·Tank가 따라오지 못하면 출하량은 늘지 않으므로 교대·인원 계획이 함께 필요합니다.
ARS는 실제 충진은 더 길 수 있지만 용기 투입·반출·퍼지를 포함하면 수동과 전체 시간이 비슷할 수 있다는 설명이 있어, 실측 전까지 수동 8 h를 그대로 적용했습니다. ARS 효과는 Capa.에 더하지 않았습니다.
충진·포장 1~1.5일, 18 h → 9 h는 검토 중인 개선 방향으로, 계산에는 반영하지 않았습니다.""")

# ============================================================ 11. 생산팀 확인 사항
sC = new_slide()
header(sC, 11, "생산팀 확인 사항", "생산팀 대화 내용과 추가 확인 항목")
sec(sC, 1.05, "① 생산·설비·품질·충진 관련 확인 내용")
cw = wb["09_확인사항"]
rows = [["구분", "생산팀 설명", "확인 필요"]]
for rr_ in range(6, 20):
    rows.append([cw.cell(row=rr_, column=c).value for c in (2, 3, 5)])
mk_table(sC, 0.31, 1.33, [1.35, 4.75, 3.27], rows, fs=7.4, left_cols=(0, 1, 2), wrap=True, rh_list=[0.27] + [0.37] * 14)
notes(sC, """생산팀과 나눈 대화 중 생산·품질·설비·충진 관련 내용만 정리했습니다. 구두로 언급된 회수량(160~170 kg, 약 220 kg, 195~200 kg)은 운전 조건과 대상이 구분되지 않아 확정 생산량으로 쓰지 않고, 환산 기준은 190 kg/Batch를 유지했습니다.
수율은 신규 Crude만이 아니라 재투입·Mix를 포함한 총 투입량 기준으로 따로 계산해야 합니다(Excel 09 시트 입력란). '135 · 120'은 단위가 확인되지 않아 생산량으로 입력하지 않았고, '65%에서 5% 상승'도 정의를 확인할 항목입니다.
이지켐은 색도 때문에 추가 투입을 제한하는 경우가 있어, 고객별 품질·색도 규격과 합격률을 별도로 확인합니다.""")

# ============================================================ 순서 정리 · 원본 3~5장 제거
order = [s1, sA, s2, sB, sV] + year_slides + [sL, sC]
lst = prs.slides._sldIdLst
ids = {prs.part.related_part(x.rId): x for x in lst}
for old in (s3, s4, s5):
    x = ids[old.part]; prs.part.drop_rel(x.rId); lst.remove(x)
for s in order:
    x = ids[s.part]; lst.remove(x); lst.append(x)
prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides")
