# -*- coding: utf-8 -*-
"""PPT 붙여넣기용 가로형 시트 (P1~P5) — 모든 값은 계산 시트·입력값에 수식으로 연결.

열 폭·글꼴·색을 PPT 표(4:3 슬라이드, 표 폭 약 9.4 in)와 맞춰, 범위를 복사해 PPT에 붙여넣으면
같은 크기의 표가 된다. build_xlsx.py에서 add_ppt_sheets(...)로 호출한다.
"""
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.properties import PageSetupProperties

PF = "Noto Sans KR"          # PPT 본문 글꼴 (SK trichem 양식)
DARK, GRAY = "1A1A1A", "7F7F7F"
HX, CXC, EZ, HS = "EB002C", "2E75B6", "FF7900", "7030A0"
line = Side(style="thin", color="D9D9D9")
BOT = Border(bottom=line)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True)
LEF = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True)


def inch_w(inches):
    """인치 → Excel 열 폭 단위 (기본 글꼴 기준 근사)."""
    px = inches * 96
    return round(px / 12, 3) if px < 12 else round((px - 5) / 7, 2)


def pfont(size=7, bold=False, color=DARK):
    return Font(name=PF, size=size, bold=bold, color=color)


def fill(c):
    return PatternFill("solid", fgColor=c)


def setup(ws, title, note, tab="404040"):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws["A1"] = title; ws["A1"].font = pfont(11, True, "EB002C")
    ws["A2"] = note; ws["A2"].font = pfont(7, False, GRAY)
    ws.column_dimensions["A"].width = 1.5


def section(ws, row, col, text, sub="", sub_col=None):
    c = ws.cell(row=row, column=col, value=text); c.font = pfont(8.5, True)
    if sub:
        ws.cell(row=row, column=sub_col or col + 4, value=sub).font = pfont(6.5, False, GRAY)
    ws.row_dimensions[row].height = 15


def put(ws, r, c, v, nf=None, bold=False, color=DARK, bg=None, align="c", size=7):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = pfont(size, bold, color); cell.border = BOT
    cell.alignment = CEN if align == "c" else LEF
    if nf:
        cell.number_format = nf
    if bg:
        cell.fill = fill(bg)
    return cell


def header_row(ws, r, c0, labels, bg="404040"):
    for j, lab in enumerate(labels):
        put(ws, r, c0 + j, lab, bold=True, color="FFFFFF", bg=bg, align="l" if j == 0 else "c")
    ws.row_dimensions[r].height = 12


KG1 = '#,##0.0;[Color10]-#,##0.0;"0"'
KG0 = '#,##0;-#,##0;"0"'
SGN = '[Color10]+#,##0;[Red]-#,##0;"0"'
SGN1 = '[Color10]+#,##0.0;[Red]-#,##0.0;"0"'
H1 = '0.0'


def add_ppt_sheets(wb, R, C, MAPX):
    M0, SR = MAPX["M0"], MAPX["SR"]
    Mq = "'04_Monthly_2026_2028'!"
    rows_04 = lambda yr, m: M0 + (yr - 2026) * 12 + (m - 1)

    # ---------------------------------------------------------------- P_공통 안내
    legend_txt = ("사용법: 표 범위를 선택 → 복사 → PPT에서 [붙여넣기 옵션: 원본 서식 유지] 또는 [그림]. "
                  "값은 01_Inputs·02~05 시트와 수식 연결 — 입력을 바꾸면 자동 갱신. 이 시트 칸에 직접 입력하면 연결이 끊김.")

    # ================================================================ P1 — As-is 공정시간
    ws = wb.create_sheet("P1_공정시간")
    setup(ws, "P1 | As-is 공정시간 · 현재 생산능력 (PPT 1장)", legend_txt, "EB002C")
    steps = ["IQC", "준비·투입", "정제", "PQC+이송", "FQC", "충진 전 소계", "충진", "OQC·출하", "확인시간 합계"]
    widths = [1.7] + [0.85] * 9
    for j, w in enumerate(widths):
        ws.column_dimensions[CL(2 + j)].width = inch_w(w)
    r = 4
    section(ws, r, 2, "① 한 Batch 확인 공정시간 (h)", "단순 합계 — 달력 납기·다음 Batch 투입 간격 아님"); r += 1
    header_row(ws, r, 2, ["경로"] + steps); r += 1
    lt0 = C["lt0"]
    for lab, col, bg in (("5 Gal 약 9병", "B", None), ("이지켐 200 L 1용기 (수동)", "C", None), ("한솔 200 L 1용기", "D", None)):
        put(ws, r, 2, lab, bold=True, align="l")
        for k in range(5):
            put(ws, r, 3 + k, f"='03_Capacity_Model'!{col}{lt0 + k}", H1)
        put(ws, r, 8, f"=SUM(C{r}:G{r})", H1, bold=True, bg="F2F2F2")
        put(ws, r, 9, f"='03_Capacity_Model'!{col}{lt0 + 5}", H1)
        put(ws, r, 10, f"='03_Capacity_Model'!{col}{lt0 + 6}", H1)
        put(ws, r, 11, f"='03_Capacity_Model'!{col}{lt0 + 8}", H1, bold=True, color=HX)
        r += 1
    put(ws, r, 2, "※ IQC 2 + 준비 2 + 정제 45 + PQC·이송 6 + FQC 2 = 57 h · 5 Gal 18 h = 2.25 근무일 · 한솔 충진시간 미확인", align="l", color=GRAY, size=6)
    ws.cell(row=r, column=2).border = Border(); r += 2
    section(ws, r, 2, "② 정제기 1대 생산능력 비교", "월 720 h 기준"); r += 1
    header_row(ws, r, 2, ["구분", "과거 최대", "", "정제 + 기타", "", "이론 (정제만)", "", "현재 Capa.", "", ""])
    for cc, ce in ((3, 4), (5, 6), (7, 8), (9, 11)):
        ws.merge_cells(start_row=r, start_column=cc, end_row=r, end_column=ce)
    r += 1
    vals = [("Batch / 월", f"={R('max11')}", "0", f"='03_Capacity_Model'!B{_row(C, 'mx_ref')}&\" h + \"&TEXT('03_Capacity_Model'!B{_row(C, 'mx_gap')},\"0\")&\" h\"", None,
             f"='03_Capacity_Model'!B{_row(C, 'th_bm')}", "0.0", f"={R('capa_now')}&\" 톤/년\"", None),
            ("kg / 월 · 산식", f"='03_Capacity_Model'!B{_row(C, 'mx_kgm')}", KG0, "합계 720 h", None,
             f"='03_Capacity_Model'!B{_row(C, 'th_kgm')}", KG0, f"=TEXT('03_Capacity_Model'!B{_row(C, 'cp_b')},\"0.0\")&\" Batch 상당 · 월 \"&TEXT('03_Capacity_Model'!B{_row(C, 'cp_bm')},\"0.00\")", None)]
    for lab, a, anf, b, bnf, c_, cnf, d_, dnf in vals:
        put(ws, r, 2, lab, bold=True, align="l")
        for col, v, nf in ((3, a, anf), (5, b, bnf), (7, c_, cnf), (9, d_, dnf)):
            put(ws, r, col, v, nf, bold=True)
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1 if col < 9 else 11)
        r += 1

    # ================================================================ P2 — 일정 · 2대 Gantt (하나의 180칸 격자: 월 = 10칸, 일 = 6칸 = 4 h)
    ws = wb.create_sheet("P2_일정_Gantt")
    setup(ws, "P2 | 투자·운영 일정 · 정제기 2대 시간차 운전 30일 (PPT 2장)", legend_txt, "EB002C")
    NC = 180
    ws.column_dimensions["B"].width = inch_w(1.3)
    gw = 7.9 / NC
    for j in range(NC):
        ws.column_dimensions[CL(3 + j)].width = inch_w(gw)
    last = CL(2 + NC)
    r = 4
    section(ws, r, 2, "① 투자 · 운영 일정 (2027.1~2028.6)", "날짜는 01_Inputs G·D 입력과 연결 — 입력 변경 시 자동 반영", sub_col=3 + 60); r += 1
    put(ws, r, 2, "구분", bold=True, color="FFFFFF", bg="404040", align="l")
    for j in range(18):
        y, m = (2027, j + 1) if j < 12 else (2028, j - 11)
        c = 3 + j * 10
        put(ws, r, c, f"=DATE({y},{m},1)", 'yy"."m', bold=True, color="FFFFFF", bg="404040" if y == 2027 else "7F7F7F", size=6)
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 9)
    hr = r; r += 1
    sched = [("■ 공사", f'IF(AND({{d}}>={R("con_s")},{{d}}<={R("con_e")}),"공사","")', "FBE5D6"),
             ("■ 시운전", f'IF(AND({{d}}>={R("trial_s")},{{d}}<{R("rf_s")}),"시운전","")', "FFF2CC"),
             ("■ 정제기 2대 적용 생산", f'IF({{d}}>={R("rf_s")},IF({{d}}<={R("f12_e")},"12개월","생산"),"")', "FDE9E7"),
             ("■ ARS 200 L 충진", f'IF({{d}}>={R("ars_hs")},"ARS","수동")', "E4DFEC"),
             ("■ 105℃ 검토안", f'IF(AND(ISNUMBER({R("t105")}),{{d}}>={R("t105")}),"105℃",IF(YEAR({{d}})=2028,"미정",""))', "E4DFEC")]
    for lab, fx, bg in sched:
        put(ws, r, 2, lab, bold=True, align="l")
        for j in range(18):
            c = 3 + j * 10
            put(ws, r, c, "=" + fx.replace("{d}", f"${CL(c)}${hr}"), size=6)
            ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 9)
        ws.conditional_formatting.add(f"C{r}:{last}{r}", FormulaRule(formula=[f'AND(C{r}<>"",C{r}<>"수동",C{r}<>"미정")'],
                                                                   fill=fill(bg), font=Font(name=PF, bold=True, color="C00000")))
        r += 1
    r += 1
    section(ws, r, 2, "② 정제기 2대 시간차 운전 개념 — 30일 (720 h)", "시차·설비별 간격 = 01_Inputs G (운영 개념 설명용 가정) · 1칸 = 4 h", sub_col=3 + 60); r += 1
    SL = 6
    put(ws, r, 2, "일 →", bold=True, color="FFFFFF", bg="404040", align="l")
    for dday in range(30):
        c = 3 + dday * SL
        put(ws, r, c, dday + 1, "0", bold=True, color="FFFFFF", bg="404040", size=6)
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + SL - 1)
    hr2 = r; r += 1
    off, gi = f"'03_Capacity_Model'!$B${_row(C, 'off')}", f"'03_Capacity_Model'!$B${_row(C, 'g_int')}"
    pre, ref = R("t_iqc"), R("t_ref")
    lanes = [("정제기 1", "0", "ref"), ("정제기 2", off, "ref"), ("PQC·이송·FQC", None, "ins"), ("5 Gal 충진 (공용)", None, "fil")]
    for lab, o, kind in lanes:
        put(ws, r, 2, lab, bold=True, align="l", size=6.5)
        for j in range(NC):
            t = f"({j}*24/{SL}+2)"
            ph = lambda oo: f"MOD({t}-{oo},{gi})"
            if kind == "ref":
                fx = f"=IF(AND({ph(o)}>={pre},{ph(o)}<{pre}+2+{ref}),1,IF({ph(o)}>={pre}+2+{ref},2,0))"
            elif kind == "ins":
                fx = f"=IF(OR(AND({ph(0)}>=49,{ph(0)}<57),AND({ph(off)}>=49,{ph(off)}<57)),3,0)"
            else:
                fx = f"=IF(OR(AND({ph(0)}>=57,{ph(0)}<75),AND({ph(off)}>=57,{ph(off)}<75)),4,0)"
            cell = ws.cell(row=r, column=3 + j, value=fx); cell.number_format = ";;;"
        ws.row_dimensions[r].height = 11
        r += 1
    rng = f"C{hr2 + 1}:{last}{r - 1}"
    for v, col in ((1, "FF7900"), (2, "FFF2CC"), (3, "2E75B6"), (4, "548235")):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[str(v)], fill=fill(col)))
    put(ws, r, 2, "누적 생산 kg", bold=True, align="l", size=6.5)
    for dday in range(30):
        c = 3 + dday * SL
        T = (dday + 1) * 24
        cnt = "+".join(f"(INT(({T}-{oo}-57)/{gi})-INT((0-{oo}-57)/{gi}))" for oo in ("0", off))
        put(ws, r, c, f"=({cnt})*{R('kg_b')}", "#,##0", color=CXC, size=5.5)
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + SL - 1)
    r += 1
    put(ws, r, 2, "범례: 주황 = 준비·정제 47 h · 연노랑 = 기타(Mix·이송·대기, 확인) · 파랑 = PQC·이송·FQC · 초록 = 5 Gal 충진 18 h", align="l", color=GRAY, size=6)
    ws.cell(row=r, column=2).border = Border()
    ws.print_area = f"B4:{last}{r}"

    # ================================================================ P3 — 과거 Batch · 2026
    ws = wb.create_sheet("P3_과거Batch_2026")
    setup(ws, "P3 | 과거 Batch 표시일자 · 간격 · 생산능력 비교 · 2026 출하 (PPT 3장)", legend_txt, "7F7F7F")
    ws.column_dimensions["B"].width = inch_w(0.55)
    for d in range(31):
        ws.column_dimensions[CL(3 + d)].width = inch_w(0.155)
    for j, w in enumerate((0.42, 0.42, 0.42)):
        ws.column_dimensions[CL(34 + j)].width = inch_w(w)
    r = 4
    section(ws, r, 2, "① Batch 표시일자 분포 (일 1~31)", "02_Batch_Raw 원자료와 수식 연결 · 빨강 = Total≠번호", sub_col=3 + 14); r += 1
    put(ws, r, 2, "월", bold=True, color="FFFFFF", bg="404040", align="l")
    for d in range(31):
        put(ws, r, 3 + d, d + 1, "0", bold=True, color="FFFFFF", bg="404040", size=5.5)
    for j, lab in enumerate(("번호", "Total", "차이")):
        put(ws, r, 34 + j, lab, bold=True, color="FFFFFF", bg="404040", size=6)
    r += 1
    Bq = "'02_Batch_Raw'!"
    D_r, E_r = f"{Bq}$D${MAPX['B0']}:$D${MAPX['BL']}", f"{Bq}$E${MAPX['B0']}:$E${MAPX['BL']}"
    d0 = r
    for m in (5, 6, 7, 8, 9, 10):
        put(ws, r, 2, f"{m}월" + (" *" if m == 5 else ""), bold=True, align="l")
        for d in range(31):
            dt_ = f"DATE({R('hist_y')},{m},{d + 1})"
            fx = f'=IF({d + 1}>DAY(EOMONTH(DATE({R("hist_y")},{m},1),0)),"",IFERROR(INDEX({D_r},MATCH({dt_},{E_r},0)),""))'
            put(ws, r, 3 + d, fx, "0", bold=True, color="FFFFFF", size=5.5)
        sr = MAPX["SUMROW"][m]
        put(ws, r, 34, f"={Bq}C{sr}", "0", bold=True)
        put(ws, r, 35, f'=IF(ISNUMBER({Bq}D{sr}),{Bq}D{sr},"–")', "0", bold=True)
        put(ws, r, 36, f'=IF(ISNUMBER({Bq}D{sr}),{Bq}E{sr},"–")', SGN, bold=True)
        r += 1
    ws.conditional_formatting.add(f"C{d0}:AG{d0}", FormulaRule(formula=[f'C{d0}<>""'], fill=fill("7F7F7F")))
    ws.conditional_formatting.add(f"C{d0 + 1}:AG{r - 1}", FormulaRule(formula=[f'C{d0 + 1}<>""'], fill=fill(EZ)))
    ws.conditional_formatting.add(f"AH{d0}:AJ{r - 1}", FormulaRule(formula=[f'AND(ISNUMBER($AJ{d0}),$AJ{d0}<>0)'], font=Font(name=PF, bold=True, color="C00000")))
    put(ws, r, 2, f"* 5월 일부(#37·#38) · 6~10월 번호 54 vs Total 46 · 표시일자 = 계획 기재일 (투입·완료 미확정)", align="l", color=GRAY, size=6)
    ws.cell(row=r, column=2).border = Border(); r += 2
    # 간격 분포
    section(ws, r, 2, "② 표시일자 간격 분포 (#39~#92)"); r += 1
    tot = MAPX["SUMROW"]["tot"]
    put(ws, r, 2, "간격", bold=True, color="FFFFFF", bg="404040", align="l")
    for j, lab in enumerate(("1일", "2일", "3일", "4일", "5일")):
        put(ws, r, 3 + j * 3, lab, bold=True, color="FFFFFF", bg="404040"); ws.merge_cells(start_row=r, start_column=3 + j * 3, end_row=r, end_column=5 + j * 3)
    put(ws, r, 18, "평균 간격", bold=True, color="FFFFFF", bg="404040"); ws.merge_cells(start_row=r, start_column=18, end_row=r, end_column=25)
    put(ws, r, 26, "중앙값", bold=True, color="FFFFFF", bg="404040"); ws.merge_cells(start_row=r, start_column=26, end_row=r, end_column=30)
    r += 1
    put(ws, r, 2, "건수", bold=True, align="l")
    for j, col in enumerate("OPQRS"):
        put(ws, r, 3 + j * 3, f"={Bq}{col}{tot}", "0", bold=True); ws.merge_cells(start_row=r, start_column=3 + j * 3, end_row=r, end_column=5 + j * 3)
    put(ws, r, 18, f"=TEXT('03_Capacity_Model'!B{_row(C, 'ob_int')},\"0.0\")&\" h (\"&TEXT('03_Capacity_Model'!B{_row(C, 'ob_lo')},\"0.0\")&\"~\"&TEXT('03_Capacity_Model'!B{_row(C, 'ob_hi')},\"0.0\")&\")\"", bold=True, color=EZ)
    ws.merge_cells(start_row=r, start_column=18, end_row=r, end_column=25)
    put(ws, r, 26, f"='03_Capacity_Model'!B{_row(C, 'ob_med')}&\" h\"", bold=True); ws.merge_cells(start_row=r, start_column=26, end_row=r, end_column=30)
    r += 2
    # 생산능력 비교 + 2026 표는 월 열 구성으로 별도 열 블록
    c0 = 38
    ws.column_dimensions[CL(c0)].width = inch_w(1.55)
    for j in range(13):
        ws.column_dimensions[CL(c0 + 1 + j)].width = inch_w(0.56)
    rr = 4
    section(ws, rr, c0, "③ 월 생산능력 비교 (정제기 1대 · 720 h)"); rr += 1
    header_row(ws, rr, c0, ["기준", "간격 h", "", "월 Batch", "", "월 kg", "", "연 환산 t", "", "26.2 대비", "", "", "", ""]); rr += 1
    for k in range(4):
        cr = MAPX["CMP0"] + k
        put(ws, rr, c0, f"='03_Capacity_Model'!A{cr}", bold=True, align="l")
        for j, (col, nf) in enumerate((("B", "0.0"), ("C", "0.0"), ("D", KG0), ("E", "0.00"), ("F", '+0.00;-0.00;"0"'))):
            put(ws, rr, c0 + 1 + j * 2, f"='03_Capacity_Model'!{col}{cr}", nf, bold=True, color=HX if k == 2 else DARK)
            ws.merge_cells(start_row=rr, start_column=c0 + 1 + j * 2, end_row=rr, end_column=c0 + 2 + j * 2)
        put(ws, rr, c0 + 11, f"='03_Capacity_Model'!H{cr}", align="l", color=GRAY, size=6)
        ws.merge_cells(start_row=rr, start_column=c0 + 11, end_row=rr, end_column=c0 + 13)
        rr += 1
    rr += 1
    section(ws, rr, c0, "⑤ 2026 출하 제시분 vs 현재 Capa. (kg)", "CXMT·한솔 미제시 — 합계 미포함"); rr += 1
    header_row(ws, rr, c0, ["구분"] + [f"{m}월" for m in range(1, 13)] + ["연간"]); rr += 1
    rows26 = [("■ SK하이닉스", "U", KG0, HX), ("■ 이지켐", "W", KG0, EZ), ("■ CXMT", "V", KG0, CXC), ("■ 한솔", "X", KG0, HS),
              ("출하 합계 (제시분)", "Y", KG0, None), ("계획 Total×190", "M", KG0, None), ("계획 번호×190", "L", KG0, None),
              ("추정 생산 가능 (관측)", "S", KG0, None), ("추정 − 출하", "AP", SGN, None)]
    for lab, col, nf, colr in rows26:
        put(ws, rr, c0, lab, bold=True, align="l", color=colr or DARK)
        for m in range(1, 13):
            src = f"{Mq}{col}{rows_04(2026, m)}"
            fx = f'=IF(ISNUMBER({src}),{src}*190,"")' if col in ("L", "M") else f"={src}"
            put(ws, rr, c0 + m, fx, nf, bold=(col in ("Y", "S", "AP")))
        a, b = rows_04(2026, 1), rows_04(2026, 12)
        if col in ("L", "M"):
            ft = f"=SUM({Mq}{col}{a}:{col}{b})*190"
        elif col in ("V", "X"):
            ft = f'=IF(COUNT({Mq}{col}{a}:{col}{b})=0,"미제시",SUM({Mq}{col}{a}:{col}{b}))'
        else:
            ft = f"=SUM({Mq}{col}{a}:{col}{b})"
        put(ws, rr, c0 + 13, ft, nf, bold=True, bg="F2F2F2")
        rr += 1
    put(ws, rr, c0, "현재 Capa. 대비", bold=True, align="l")
    put(ws, rr, c0 + 1, f"=TEXT('05_Reflux_Scenarios'!B{MAPX['CMPY']['26_now']},\"0.00\")&\"톤 = 26.2톤의 \"&TEXT('05_Reflux_Scenarios'!B{MAPX['CMPY']['26_now']}/{R('capa_now')},\"0.0%\")&\" · 산술 차이 \"&TEXT('05_Reflux_Scenarios'!D{MAPX['CMPY']['26_now']},\"0.00\")&\"톤 (가동률·판매 가능량 아님)\"",
        bold=True, color=HX, align="l")
    ws.merge_cells(start_row=rr, start_column=c0 + 1, end_row=rr, end_column=c0 + 13)

    # ================================================================ P4 / P5 — 2027 · 2028 월별
    for yr, name, slide in ((2027, "P4_2027_월별", 4), (2028, "P5_2028_월별", 5)):
        ws = wb.create_sheet(name)
        setup(ws, f"P{slide} | {yr} 월별 출하 · 참고 Capa. · 충진 (PPT {slide}장 표) — kg",
              legend_txt + " 출하량 수정은 01_Inputs K(기준값) 또는 하단 출하 표에서.", "FF7900")
        ws.column_dimensions["B"].width = inch_w(2.1)
        for j in range(12):
            ws.column_dimensions[CL(3 + j)].width = inch_w(0.53)
        ws.column_dimensions["O"].width = inch_w(0.85)
        r = 4
        header_row(ws, r, 2, ["구분"] + [f"{m}월" for m in range(1, 13)] + ["연간"]); r += 1
        srr = SR[yr]
        a, b = rows_04(yr, 1), rows_04(yr, 12)
        body = [("적용 조건", "D", None, None, "-"),
                ("참고 Capa. 월 환산 (연간 환산÷12)", "T", KG1, None, f"=SUM({Mq}T{a}:T{b})"),
                ("실제 생산 가능량 (조건 확인)", "S", KG0, None, f"={Mq}S{srr}")]
        if yr == 2028:
            body.append(("105℃ 검토안 50.2 (÷12)", None, KG1, None, f"={R('capa_t')}*1000"))
        body += [("참고 Capa. − 출하", "AO", SGN, None, f"=SUM({Mq}AO{a}:AO{b})"),
                 ("■ SK하이닉스", "U", KG1, HX, f"=SUM({Mq}U{a}:U{b})"),
                 ("■ CXMT", "V", KG0, CXC, f"=SUM({Mq}V{a}:V{b})"),
                 ("■ 이지켐", "W", KG0, EZ, f"=SUM({Mq}W{a}:W{b})"),
                 ("■ 한솔", "X", KG0, HS, f"=SUM({Mq}X{a}:X{b})"),
                 ("출하 합계", "Y", KG1, None, f"=SUM({Mq}Y{a}:Y{b})"),
                 ("필요 Batch (출하÷190)", "Z", "0.0", None, f"=SUM({Mq}Z{a}:Z{b})"),
                 ("5 Gal 충진 h (병 상당×2)", "AB", "0.0", None, f"=SUM({Mq}AB{a}:AB{b})"),
                 ("이지켐 용기 (140 kg)", "AD", "0", None, f"=SUM({Mq}AD{a}:AD{b})"),
                 ("한솔 용기 (150 kg)", "AE", "0", None, f"=SUM({Mq}AE{a}:AE{b})"),
                 ("200 L 충진 방식 (이지켐/한솔)", "AFAG", None, None, "-"),
                 ("200 L 충진 h (확인분+미확인)", "AJX", "0.0", None, "-"),
                 ("기말재고", "AZ", KG0, None, "-")]
        for lab, col, nf, colr, total in body:
            put(ws, r, 2, lab, bold=True, align="l", color=colr or DARK)
            for m in range(1, 13):
                rr_ = rows_04(yr, m)
                if col is None:
                    fx = f"={R('capa_t')}*1000/12"
                elif col == "AFAG":
                    fx = f"={Mq}AF{rr_}&\"/\"&{Mq}AG{rr_}"
                elif col == "AJX":
                    fx = f'=IF(ISNUMBER({Mq}AJ{rr_}),{Mq}AJ{rr_},TEXT(N({Mq}AH{rr_})+N({Mq}AI{rr_}),"0")&"+미확인")'
                else:
                    fx = f"={Mq}{col}{rr_}"
                put(ws, r, 2 + m, fx, nf, bold=col in ("Y", "Z", "AB", "AO", "S"), size=6.5 if col in ("D", "AFAG", "AJX", "S", "AZ") else 7)
            put(ws, r, 15, total, {KG1: KG0, "0.0": "#,##0.0"}.get(nf, nf), bold=True, bg="F2F2F2")
            if col in ("D",):
                ws.conditional_formatting.add(f"C{r}:N{r}", FormulaRule(formula=[f'LEFT(C{r},2)="개선"'], fill=fill("FDE9E7")))
                ws.conditional_formatting.add(f"C{r}:N{r}", FormulaRule(formula=[f'OR(C{r}="공사",C{r}="시운전")'], fill=fill("FBE5D6")))
            if col in ("S", "AJX", "AZ"):
                ws.conditional_formatting.add(f"C{r}:O{r}", FormulaRule(formula=[f'NOT(ISNUMBER(C{r}))'], fill=fill("FFF2CC"), font=Font(name=PF, color="C55A11", size=6.5)))
            r += 1
        cy = MAPX["CMPY"]
        if yr == 2027:
            lines = [f"=\"상반기 출하 \"&TEXT('05_Reflux_Scenarios'!B{cy['27_h1']},\"0.00\")&\"톤 vs 26.2×6/12 = \"&TEXT('05_Reflux_Scenarios'!C{cy['27_h1']},\"0.0\")&\"톤 (참고) · 하반기 출하 \"&TEXT('05_Reflux_Scenarios'!B{cy['27_h2']},\"0.00\")&\"톤 vs 47.2×6/12 = \"&TEXT('05_Reflux_Scenarios'!C{cy['27_h2']},\"0.0\")&\"톤 → \"&TEXT('05_Reflux_Scenarios'!D{cy['27_h2']},\"+0.00;-0.00\")&\"톤\"",
                     f"=\"연간 출하 \"&TEXT('05_Reflux_Scenarios'!B{cy['27_now']},\"0.00\")&\"톤 · 현재 26.2 대비 \"&TEXT('05_Reflux_Scenarios'!D{cy['27_now']},\"+0.00;-0.00\")&\"톤 · 개선 47.2 대비 \"&TEXT('05_Reflux_Scenarios'!D{cy['27_rf']},\"+0.00;-0.00\")&\"톤 (연간 환산 — 실제 여유·부족 아님)\""]
        else:
            lines = [f"=\"출하 가정 \"&TEXT('05_Reflux_Scenarios'!B{cy['28_base']},\"0.00\")&\"톤 · 47.2 기준 \"&TEXT('05_Reflux_Scenarios'!D{cy['28_base']},\"+0.00;-0.00\")&\"톤 · 105℃ 50.2 적용 시 \"&TEXT('05_Reflux_Scenarios'!D{cy['28_105']},\"+0.00;-0.00\")&\"톤 (검토안·미확정 · 연간 환산 비교)\""]
        for fx in lines:
            put(ws, r, 2, "비교 (t)", bold=True, align="l")
            put(ws, r, 3, fx, bold=True, color=HX, align="l"); ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=15)
            r += 1
        put(ws, r, 2, "* 노란 칸 = 미확인 조건 (01_Inputs 입력 시 자동 계산) · 하이닉스 = 22,720÷12 (표시 반올림) · 5 Gal 병 상당 = (하이닉스+CXMT)÷20 · 참고 Capa.는 실제 생산량 아님",
            align="l", color=GRAY, size=6)
        ws.cell(row=r, column=2).border = Border()
        ws.freeze_panes = "C5"
        ws.print_area = f"B4:O{r}"
    for ws in wb.worksheets:
        if ws.title.startswith("P1") or ws.title.startswith("P3"):
            ws.print_area = f"B4:{CL(ws.max_column)}{ws.max_row}"


def _row(C, key):
    return int(C[key].split("$")[-1])
