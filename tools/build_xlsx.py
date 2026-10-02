# -*- coding: utf-8 -*-
"""SKTC CpZr 생산능력 계산 Excel (v3).

구성: 01_입력 한 곳에서만 값을 고치면 02~10 시트가 모두 따라 바뀐다 (이름 정의 사용).
사용: python build_xlsx.py <out.xlsx>   → out.xlsx + out.xlsx.map.json (PPT 스크립트가 읽는 셀 주소)
"""
import datetime as dt
import json
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

OUT = sys.argv[1]
FN = "맑은 고딕"
BLUE, BLACK, GRAY, RED, WHITE = "0000FF", "000000", "808080", "C00000", "FFFFFF"
F_IN, F_ASK, F_HDR, F_SEC, F_KEY, F_SUB = "DDEBF7", "FFF2CC", "404040", "F2F2F2", "FDE9E7", "D9D9D9"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
NF = {"kg": '#,##0;[Red]-#,##0;"-"', "kg1": '#,##0.0;[Red]-#,##0.0;"-"', "h": '#,##0.0', "h0": '0', "b": '0.0', "t": '0.00',
      "pct": '0%', "d": 'yyyy-mm-dd', "m": 'm"월"', "ym": 'yy"."m', "int": '0', "sg": '+#,##0;[Red]-#,##0;0'}
MONTHS = list(range(1, 13))
YEARS = (2026, 2027, 2028)

wb = Workbook()
REF = {}      # PPT용 셀 주소 (key -> "시트!셀")


def font(color=BLACK, bold=False, size=10, italic=False):
    return Font(name=FN, color=color, bold=bold, size=size, italic=italic)


def fill(c):
    return PatternFill("solid", fgColor=c)


def put(ws, r, c, v, kind="calc", nf=None, bold=False, align=None, wrap=False):
    """kind: in(입력) · ask(확인 필요 입력) · calc(수식) · key(핵심 결과) · lab(항목명) · hdr(머리글) · sub(소계 머리) · note(설명)"""
    x = ws.cell(row=r, column=c, value=v)
    x.border = BOX
    if kind == "in":
        x.font = font(BLUE, True); x.fill = fill(F_IN)
    elif kind == "ask":
        x.font = font(BLUE, True); x.fill = fill(F_ASK)
    elif kind == "key":
        x.font = font(BLACK, True); x.fill = fill(F_KEY)
    elif kind == "lab":
        x.font = font(BLACK, True)
    elif kind == "hdr":
        x.font = font(WHITE, True); x.fill = fill(F_HDR)
    elif kind == "sub":
        x.font = font(BLACK, True); x.fill = fill(F_SEC)
    elif kind == "note":
        x.font = font(GRAY, size=9); x.border = Border()
    else:
        x.font = font(BLACK, bold)
    if nf:
        x.number_format = NF.get(nf, nf)
    h = align or ("left" if kind in ("lab", "note", "sub") else "center")
    x.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    return x


def title(ws, text, sub):
    ws.sheet_view.showGridLines = False
    ws["B1"] = text; ws["B1"].font = font("EB002C", True, 14)
    ws["B2"] = sub; ws["B2"].font = font(GRAY, size=9)
    ws.column_dimensions["A"].width = 2


def section(ws, r, text, c=2):
    ws.cell(row=r, column=c, value=text).font = font("EB002C", True, 11)
    ws.row_dimensions[r].height = 20


def widths(ws, spec):
    for k, v in spec.items():
        ws.column_dimensions[k].width = v


def nm(name, ws, r, c):
    ref = f"'{ws.title}'!${CL(c)}${r}"
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
    REF[name] = (ws.title, f"{CL(c)}{r}")
    return name


def q(ws):
    return f"'{ws.title}'!"


def month_hdr(ws, r, yr, c0=3, label="구분", total="연간"):
    put(ws, r, c0 - 1, label, "hdr", align="left")
    for m in MONTHS:
        put(ws, r, c0 - 1 + m, dt.date(yr, m, 1), "hdr", NF["m"])
    put(ws, r, c0 + 12, total, "hdr")


# =============================================================== 01_입력
wsI = wb.active; wsI.title = "01_입력"
title(wsI, "01_입력 | 모든 기준값 (이 시트만 수정)",
      "파란 글자 = 입력 · 노란 칸 = 확인 후 입력 · 검정 = 자동 계산. 다른 시트는 모두 이 시트를 참조합니다.")
widths(wsI, {"B": 34, "C": 13, "D": 10, "E": 58})
for j, (k, lab) in enumerate((("in", "입력"), ("ask", "확인 후 입력"), ("calc", "자동 계산"), ("key", "주요 결과"))):
    put(wsI, 3, 6 + 2 * j, "", k); wsI.cell(row=3, column=7 + 2 * j, value=lab).font = font(GRAY, size=9)
r = 5


def head(text):
    global r
    r += 1
    section(wsI, r, text); r += 1
    for j, lab in enumerate(("항목", "값", "단위", "설명")):
        put(wsI, r, 2 + j, lab, "hdr", align="left" if j in (0, 3) else None)
    r += 1


def row(name, label, val, unit, note="", kind="in", nf="h"):
    global r
    if val is None and kind == "in":
        kind = "ask"
    put(wsI, r, 2, label, "lab")
    put(wsI, r, 3, val, kind, nf)
    put(wsI, r, 4, unit)
    put(wsI, r, 5, note, "note"); wsI.cell(row=r, column=5).border = BOX
    if name:
        nm(name, wsI, r, 3)
    r += 1


head("A. 공정시간 (IQC~FQC 57 h 구성)")
row("T_57", "IQC~FQC 합계 (현재 보고 기준)", 57, "h/Batch", "충진·OQC 미포함")
row("T_IQC", "수입검사 IQC", 2, "h", "원료 Lot 단위 — 여러 Batch 공통이면 Batch마다 중복 배정 안 함")
row("T_PREP", "준비·투입", 2, "h")
row("T_PQC", "공정검사 PQC ~ 제품 이송 (합계)", 6, "h", "두 공정 합계 6 h (각각 6 h 아님)")
row("T_FQC", "제품검사 FQC", 2, "h")
row("T_REF", "순수 정제 운전 (차감 계산)", "=T_57-T_IQC-T_PREP-T_PQC-T_FQC", "h", "57 − 2 − 2 − 6 − 2 = 45 h · 실측값 아님", "key")
head("B. 정제기 운전 · 점유")
row("N_NOW", "현재 정제기 수", 1, "대", "", nf="int")
row("N_NEW", "리플럭스 이후 정제기 수", 2, "대", "투입 시점을 엇갈리게 병행 운전", nf="int")
row("MH", "월 비교 기준 시간", 720, "h", "30일 × 24 h", nf="h0")
row("MAX_B", "과거 월 최대 Batch (정제기 1대)", 11, "Batch/월", "", nf="int")
row("OFFSET_H", "정제기 2 투입 시차 (운전 예시용)", 30, "h", "투입 간격 미확정 — 초류 진행에 연계 (현장 설명)")
row("TOBE_M", "정제기 2대 운영 예시 월", dt.date(2027, 7, 1), "", "운영 마일스톤 달력", nf="d")
row("OCC_ADD", "정제기 추가 점유 (배출·세척·전환)", None, "h/Batch", "정제기 해제 시점 확인")
row("INT_NEXT", "다음 Batch 투입 간격 (실측)", None, "h", "공란이면 운전 예시는 Capa. 역산값 사용")
row("T_WAIT", "검사·승인 대기", None, "h/Batch")
row("T_MIX", "Mix·Premix 준비 (설비 사용)", None, "h/Batch", "47.2톤이 현재의 2배가 아닌 이유 중 하나")
head("C. 충진 · OQC")
row("FILL_5G", "5 Gal 충진 (글로브 박스)", 2, "h/병", "하이닉스·CXMT")
row("BOT_B", "5 Gal 병 수 / Batch", 9, "병", "9병 × 20 kg = 180 kg", nf="int")
row("KG_BOT", "5 Gal 충진량", 20, "kg/병", "", nf="int")
row("OQC_5G", "5 Gal OQC·출하 작업", 2, "h/9병")
row("KG_EZ", "이지켐 200 L 충진량", 140, "kg/용기", "", nf="int")
row("FILL_EZ", "이지켐 200 L 수동 충진", 8, "h/용기")
row("KG_HS", "한솔 200 L 충진량", 150, "kg/용기", "", nf="int")
row("FILL_HS", "한솔 200 L 수동 충진", 8, "h/용기")
row("OQC_200", "200 L OQC·출하 작업", 2, "h/용기")
row("ARS_EZ_S", "이지켐 ARS 시작", dt.date(2027, 7, 1), "", "", nf="d")
row("ARS_HS_S", "한솔 ARS 시작", dt.date(2027, 7, 1), "", "'27.1~6 수동", nf="d")
ars_row = r
row("ARS_EZ_T", "이지켐 ARS 충진 (아래 단계 합계)", None, "h/용기", "공란이면 수동 8 h로 계산", "calc")
row("ARS_HS_T", "한솔 ARS 충진 (아래 단계 합계)", None, "h/용기", "공란이면 수동 8 h로 계산", "calc")
row("WH_DAY", "글로브 박스 근무시간", 8, "h/일", "1교대 기준")
row("WD_MON", "월 근무일", 22, "일/월", "", nf="int")
row("FILL_5G_EL", "5 Gal 9병 충진 경과 (근무시간 반영)", 48, "h", "현장 설명 약 2일 (작업 18 h)")
r += 1
put(wsI, r, 2, "ARS 작업 단계 (h/용기)", "hdr", align="left"); put(wsI, r, 3, "이지켐", "hdr"); put(wsI, r, 4, "한솔", "hdr")
put(wsI, r, 5, "수동과 ARS의 전체 시간 차이 확인용", "hdr", align="left"); r += 1
st0 = r
for s in ("용기 준비", "투입·반출", "퍼지", "실제 충진", "설비 전환", "검사·OQC", "작업자 투입시간", "설비 점유시간"):
    put(wsI, r, 2, s, "lab"); put(wsI, r, 3, None, "ask", "h"); put(wsI, r, 4, None, "ask", "h"); put(wsI, r, 5, "", "note"); r += 1
st1 = r - 1

a, b = st0, st1 - 2
wsI.cell(row=ars_row, column=3).value = f'=IF(COUNT(C{a}:C{b})=0,"",SUM(C{a}:C{b}))'
wsI.cell(row=ars_row + 1, column=3).value = f'=IF(COUNT(D{a}:D{b})=0,"",SUM(D{a}:D{b}))'
head("D. Batch 생산량")
row("KG_B", "Batch당 생산량 (환산 기준)", 190, "kg/Batch", "생산팀 언급 회수량은 09_확인사항에 별도 보존", nf="int")
head("E. Capa. 기준")
row("CAPA_NOW", "현재 연간 Capa.", 26.2, "t/년", "", nf="t")
row("CAPA_ADD", "리플럭스 증가분", 21, "t/년", "", nf="t")
row("CAPA_RF", "개선 후 연간 Capa.", "=CAPA_NOW+CAPA_ADD", "t/년", "현재의 2배(52.4) 아님 — Mix·Premix 준비 등", "key", "t")
row("CAPA_TADD", "105℃ 증가분 (2028)", 3, "t/년", "열 안정성·Dimer·Unknown·수율 검증 필요", nf="t")
row("CAPA_105", "2028 Capa. (105℃)", "=CAPA_RF+CAPA_TADD", "t/년", "", "key", "t")
row("MAINT_M", "대정비 월 (매년)", 4, "월", "", nf="int")
row("RF_START", "정제기 2대 생산 시작", dt.date(2027, 5, 1), "", "대정비 다음 달", nf="d")
row("T_105", "105℃ 적용 시작", dt.date(2028, 1, 1), "", "", nf="d")
maint_row = r
row("MAINT_KG", "대정비 월 생산량", None, "kg", "아래 2027년 4월 생산계획과 연결", "calc", "kg")
row("CAPA_M", "월 Capa. (47.2톤 ÷ 11)", "=CAPA_RF*1000/11", "kg/월", "47,200 ÷ 11 · 대정비 월은 별도 (1,188 kg)", "key", "kg1")
row("CAPA_M105", "월 Capa. (50.2톤 ÷ 11)", "=CAPA_105*1000/11", "kg/월", "50,200 ÷ 11 · 대정비 월은 별도 (1,188 kg)", "key", "kg1")
head("F. 재고 · 품질")
row("INV_0", "기초재고 (2027-01-01)", 2882, "kg", "재고 = 전월 재고 + 생산 − 출하", nf="kg")
row("YIELD_Q", "양품률 (입력 시 재고에 반영)", None, "%", "공란이면 생산계획 그대로 재고 계산", nf="pct")

# ---- 월별 입력: 생산계획 · 출하계획
r += 1
section(wsI, r, "G. 월별 생산계획 (kg) — 2027년 1~4월은 확정치로 두 버전 공통"); r += 1
PLAN = {}
PLAN_B = {2027: [2178, 2178, 3168, 1188] + [3564] * 8,
          2028: [3564, 3564, 3564, 1188, 3564, 3564, 3564, 3168, 3168, 3168, 3168, 3168]}
for yr in (2027, 2028):
    month_hdr(wsI, r, yr, total=f"{yr} 합계"); r += 1
    put(wsI, r, 2, "기존 생산계획" + (" (1~4월 확정)" if yr == 2027 else ""), "lab")
    for m in MONTHS:
        put(wsI, r, 2 + m, PLAN_B[yr][m - 1], "in", "kg")
    put(wsI, r, 15, f"=SUM(C{r}:N{r})", "key", "kg")
    PLAN[yr] = r; r += 2
wsI.cell(row=maint_row, column=3).value = f"=INDEX(C{PLAN[2027]}:N{PLAN[2027]},MAINT_M)"
section(wsI, r, "H. 고객별 월 출하계획 (kg) — 2028년은 비교용 가정"); r += 1
SHIP = {}
SHIP_DEF = {
    2026: {"hx": [1600, 1420, 1440, 1580, 1200, 1200, 1360, 1420, 1380, 1480, 1480, 1420], "cx": [None] * 12,
           "ez": [0] * 8 + [560] * 4, "hs": [None] * 12},
    2027: {"hx": ["=22720/12"] * 12, "cx": [0, 0, 580, 580, 580, 580] + [780] * 6, "ez": [560] * 6 + [840] * 6, "hs": [300] * 12},
    2028: {"hx": ["=22720/12"] * 12, "cx": [780] * 12, "ez": [840] * 12, "hs": [300] * 12}}
CUST = (("hx", "SK하이닉스"), ("cx", "CXMT"), ("ez", "이지켐"), ("hs", "한솔"))
for yr in YEARS:
    month_hdr(wsI, r, yr, label=f"{yr}", total="합계"); r += 1
    for k, lab in CUST:
        put(wsI, r, 2, lab, "lab")
        for m in MONTHS:
            v = SHIP_DEF[yr][k][m - 1]
            put(wsI, r, 2 + m, v, "in" if v is not None else "ask", "kg1" if k == "hx" else "kg")
        put(wsI, r, 15, f'=IF(COUNT(C{r}:N{r})=0,"미제시",SUM(C{r}:N{r}))', "calc", "kg")
        SHIP[(yr, k)] = r; r += 1
    put(wsI, r, 2, "합계", "sub")
    for m in range(13):
        col = CL(3 + m)
        put(wsI, r, 3 + m, f"=SUM({col}{r - 4}:{col}{r - 1})", "key", "kg")
    SHIP[(yr, "tot")] = r; r += 2
put(wsI, r, 2, "2026 CXMT·한솔 미제시 칸은 0이 아니라 공란 · 하이닉스 22,720 ÷ 12 균등 · 이지켐 2027.7~ 840 kg/월", "note"); r += 2
section(wsI, r, "I. 생산계획 Batch 완료일 — 완료일 = 해당 Batch의 생산 및 충진 완료 예정일 (Batch마다 독립)"); r += 1
for j, lab in enumerate(("Batch No.", "완료일", "비고")):
    put(wsI, r, 2 + j, lab, "hdr", align="left" if j != 1 else None)
r += 1
PLAN_DT0 = r
for n, d in [(81, dt.date(2026, 9, 28))] + [(82 + i, dt.date(2026, 10, x)) for i, x in enumerate((1, 4, 7, 10, 12, 15, 18, 20, 23, 26, 29))]:
    put(wsI, r, 2, f"#{n}", "in", align="left"); put(wsI, r, 3, d, "in", "d"); put(wsI, r, 4, "9월 마지막 Batch (간격 기준점)" if n == 81 else "", "note")
    r += 1
PLAN_DT1 = r - 1
put(wsI, r, 2, "10월 완료 #82~#92 = 11 Batch · 행을 추가하면 03_Batch주기 계산 범위도 같이 조정", "note"); r += 1
wsI.freeze_panes = "C5"
Iq = q(wsI)


def ship(yr, k, m):
    return f"{Iq}${CL(2 + m)}${SHIP[(yr, k)]}"


def plan_b(yr, m):
    return f"{Iq}${CL(2 + m)}${PLAN[yr]}"


# =============================================================== 02_공정시간
ws = wb.create_sheet("02_공정시간")
title(ws, "02_공정시간 | 57 h 구성 · 경로별 단순 합계 · 시간 구분", "PPT 1장 표. 단순 합계는 실제 달력 Lead Time이나 다음 Batch 투입 간격이 아닙니다.")
widths(ws, {"B": 30, **{CL(i): 12 for i in range(3, 12)}, "L": 40})
section(ws, 4, "① 57 h 구성 (IQC~FQC)")
for j, lab in enumerate(("IQC", "준비·투입", "순수 정제", "PQC~이송", "FQC", "합계")):
    put(ws, 5, 3 + j, lab, "hdr")
put(ws, 5, 2, "구분", "hdr", align="left")
put(ws, 6, 2, "시간 (h)", "lab")
for j, k in enumerate(("T_IQC", "T_PREP", "T_REF", "T_PQC", "T_FQC")):
    put(ws, 6, 3 + j, f"={k}", "key" if k == "T_REF" else "calc", "h")
put(ws, 6, 8, "=SUM(C6:G6)", "key", "h")
put(ws, 7, 2, "근거", "lab")
for j, t in enumerate(("제공", "제공", "57 − 12 차감", "합계 6 h", "제공", "보고 기준")):
    put(ws, 7, 3 + j, t)
section(ws, 9, "② 경로별 공정시간 단순 합계 (h)")
hd = ("IQC~FQC", "충진", "OQC·출하", "합계", "충진량 (kg)", "충진 기준")
put(ws, 10, 2, "경로", "hdr", align="left")
for j, lab in enumerate(hd):
    put(ws, 10, 3 + j, lab, "hdr")
paths = (("5 Gal 9병 (하이닉스·CXMT)", "=FILL_5G*BOT_B", "=OQC_5G", "=BOT_B*KG_BOT", "2 h/병 × 9병"),
         ("이지켐 200 L 1용기 (수동)", "=FILL_EZ", "=OQC_200", "=KG_EZ", "8 h/용기"),
         ("한솔 200 L 1용기 (수동)", "=FILL_HS", "=OQC_200", "=KG_HS", "8 h/용기"))
for i, (lab, fil, oqc, kg, basis) in enumerate(paths):
    rr = 11 + i
    put(ws, rr, 2, lab, "lab")
    put(ws, rr, 3, "=T_57", "calc", "h"); put(ws, rr, 4, fil, "calc", "h"); put(ws, rr, 5, oqc, "calc", "h")
    put(ws, rr, 6, f"=C{rr}+D{rr}+E{rr}", "key", "h"); put(ws, rr, 7, kg, "calc", "kg"); put(ws, rr, 8, basis)
    REF[f"lt{i}"] = (ws.title, f"F{rr}"); REF[f"fill{i}"] = (ws.title, f"D{rr}")
section(ws, 15, "③ 시간 구분 — Lead Time · 정제기 점유 · 충진 작업")
put(ws, 16, 2, "항목", "hdr", align="left"); put(ws, 16, 3, "값", "hdr"); put(ws, 16, 4, "단위", "hdr")
put(ws, 16, 5, "비고", "hdr", align="left"); ws.merge_cells("E16:I16")
items = (("정제기 점유 확인분 (준비·투입 + 순수 정제)", "=T_PREP+T_REF", "h", "추가 점유 (배출·세척·전환) 별도 확인"),
         ("정제기 추가 점유", '=IF(ISNUMBER(OCC_ADD),OCC_ADD,"확인 필요")', "h", "01_입력 B"),
         ("다음 Batch 투입 간격 (실측)", '=IF(ISNUMBER(INT_NEXT),INT_NEXT,"확인 필요")', "h", "과거 최대 11 Batch 기준 720 ÷ 11 = 65.5 h"),
         ("검사·승인 대기", '=IF(ISNUMBER(T_WAIT),T_WAIT,"확인 필요")', "h", ""),
         ("5 Gal 9병 충진 작업", "=FILL_5G*BOT_B", "h", "실제 작업시간"),
         ("5 Gal 9병 충진 경과 (근무시간 반영)", "=FILL_5G*BOT_B/WH_DAY", "근무일", "현장 설명 약 2일 · 교대·연장근무 조건 확인"),
         ("190 kg − 5 Gal 9병 180 kg", "=KG_B-BOT_B*KG_BOT", "kg", "잔량·추가 충진·다른 고객 배분 확인 (손실 처리 안 함)"))
for i, (lab, fx, unit, note) in enumerate(items):
    rr = 17 + i
    put(ws, rr, 2, lab, "lab"); put(ws, rr, 3, fx, "calc", "0.0"); put(ws, rr, 4, unit)
    put(ws, rr, 5, note, "note"); ws.merge_cells(f"E{rr}:I{rr}")
REF["occ"] = (ws.title, "C17"); REF["el_5g"] = (ws.title, "C22")

# =============================================================== 04_Capa기준 (03보다 먼저 정의되는 값 없음 — 순서만 뒤)
wsC = wb.create_sheet("03_Batch주기")
wsK = wb.create_sheet("04_Capa기준")

# ---- 03_Batch주기
ws = wsC
title(ws, "03_Batch주기 | 생산계획 완료 주기 · 순수 작업 vs 대기 · 운영 마일스톤", "완료일 = 생산 및 충진 완료. 정제기 주기 = 순수 점유(준비·투입 + 순수 정제) + 대기.")
widths(ws, {"B": 30, **{CL(i): 12 for i in range(3, 12)}})
section(ws, 4, "① 생산계획 Batch 완료 주기")
for j, lab in enumerate(("Batch", "완료일", "요일", "간격 (일)", "간격 (h)", "10월 완료", "10월 누적 kg")):
    put(ws, 5, 2 + j, lab, "hdr")
n_pl = PLAN_DT1 - PLAN_DT0 + 1
for i in range(n_pl):
    rr = 6 + i; src = PLAN_DT0 + i
    put(ws, rr, 2, f"={Iq}B{src}", "calc"); put(ws, rr, 3, f"={Iq}C{src}", "calc", "d")
    put(ws, rr, 4, f'=CHOOSE(WEEKDAY(C{rr},2),"월","화","수","목","금","토","일")', "calc")
    put(ws, rr, 5, "" if i == 0 else f"=C{rr}-C{rr - 1}", "calc", "int"); put(ws, rr, 6, "" if i == 0 else f"=E{rr}*24", "calc", "h0")
    put(ws, rr, 7, f"=IF(MONTH(C{rr})=10,1,0)", "calc", "int"); put(ws, rr, 8, f"=SUM($G$6:G{rr})*KG_B", "calc", "kg")
e = 5 + n_pl
REF["pl0"] = (ws.title, "B6"); REF["pl_n"] = (ws.title, f"B{e}")
stats = (("10월 완료 Batch", f"=SUM(G6:G{e})", "int"), ("10월 생산량 (kg)", f"=C{e + 2}*KG_B", "kg"),
         ("평균 완료 간격 (h)", f"=(C{e}-C6)*24/{n_pl - 1}", "h"), ("평균 완료 간격 (일)", f"=C{e + 4}/24", "b"),
         ("3일 간격 (회)", f"=COUNTIF(E7:E{e},3)", "int"), ("2일 간격 (회)", f"=COUNTIF(E7:E{e},2)", "int"))
for i, (lab, fx, nf) in enumerate(stats):
    rr = e + 2 + i
    put(ws, rr, 2, lab, "lab"); put(ws, rr, 3, fx, "key", nf)
nm("PL_B", ws, e + 2, 3); nm("PL_KG", ws, e + 3, 3); nm("CYC_ASIS", ws, e + 4, 3); nm("PL_D3", ws, e + 6, 3); nm("PL_D2", ws, e + 7, 3)
r = e + 10
section(ws, r, "② 순수 작업 vs 대기 (평균)"); r += 1
for j, lab in enumerate(("구분", "순수 작업 (h)", "대기 (h)", "실제 소요 (h)", "일 환산", "순수 비율", "산출")):
    put(ws, r, 2 + j, lab, "hdr", align="left" if j in (0, 6) else None)
r += 1
WAIT0 = r
wt = (("정제기 Batch 주기", "=T_PREP+T_REF", "=CYC_ASIS-C{r}", "=C{r}+D{r}", "완료 간격 − (준비·투입 + 순수 정제)"),
      ("5 Gal 1 Batch (IQC~OQC)", "=T_57+FILL_5G*BOT_B+OQC_5G", "=WAIT_ASIS", "=C{r}+D{r}", "순수 77 h + 평균 대기"),
      ("200 L 1 Batch (IQC~OQC)", "=T_57+FILL_EZ+OQC_200", "=WAIT_ASIS", "=C{r}+D{r}", "순수 67 h + 평균 대기"),
      ("5 Gal 9병 충진 (근무시간)", "=FILL_5G*BOT_B", "=FILL_5G_EL-C{r}", "=FILL_5G_EL", "1교대 · 현장 설명 약 2일"))
for i, (lab, pu, wa, ac, basis) in enumerate(wt):
    rr = r + i
    put(ws, rr, 2, lab, "lab")
    put(ws, rr, 3, pu, "calc", "h"); put(ws, rr, 4, wa.replace("{r}", str(rr)), "key", "h"); put(ws, rr, 5, ac.replace("{r}", str(rr)), "key", "h")
    put(ws, rr, 6, f"=E{rr}/24", "calc", "b"); put(ws, rr, 7, f"=C{rr}/E{rr}", "calc", "pct"); put(ws, rr, 8, basis, "note")
    ws.cell(row=rr, column=8).border = BOX
    REF[f"w{i}"] = (ws.title, f"B{rr}")
nm("WAIT_ASIS", ws, r, 4); nm("LT_5G", ws, r + 1, 5); nm("LT_2L", ws, r + 2, 5)
r += len(wt) + 2
section(ws, r, "③ 평균 Batch 마일스톤 (대기 포함, h)"); r += 1
for j, lab in enumerate(("단계", "5 Gal 시간", "시작", "종료", "200 L 시간", "시작", "종료")):
    put(ws, r, 2 + j, lab, "hdr", align="left" if j == 0 else None)
r += 1
MS0 = r
steps = (("IQC", "T_IQC", "T_IQC"), ("준비·투입", "T_PREP", "T_PREP"), ("순수 정제", "T_REF", "T_REF"), ("PQC~이송", "T_PQC", "T_PQC"),
         ("FQC", "T_FQC", "T_FQC"), ("대기 (평균)", "WAIT_ASIS", "WAIT_ASIS"), ("충진", "FILL_5G*BOT_B", "FILL_EZ"), ("OQC·출하", "OQC_5G", "OQC_200"))
for i, (lab, a5, a2) in enumerate(steps):
    rr = r + i
    put(ws, rr, 2, lab, "lab")
    put(ws, rr, 3, f"={a5}", "calc", "h"); put(ws, rr, 4, "=0" if i == 0 else f"=E{rr - 1}", "calc", "h"); put(ws, rr, 5, f"=D{rr}+C{rr}", "calc", "h")
    put(ws, rr, 6, f"={a2}", "calc", "h"); put(ws, rr, 7, "=0" if i == 0 else f"=H{rr - 1}", "calc", "h"); put(ws, rr, 8, f"=G{rr}+F{rr}", "calc", "h")
REF["ms0"] = (ws.title, f"B{MS0}")
r += len(steps) + 1
section(ws, r, "④ 월 생산량 비교 — 설비별 주기 = 순수 점유 + 대기"); r += 1
for j, lab in enumerate(("기준", "정제기", "Batch/월", "kg/월", "연간 (t)", "설비별 주기 (h)", "순수 (h)", "대기 (h)")):
    put(ws, r, 2 + j, lab, "hdr", align="left" if j == 0 else None)
r += 1
CMP0 = r
cmp_rows = (("이론 (순수 정제만)", "=N_NOW", "=INT(MH/T_REF)", "=D{r}*KG_B", "=E{r}*12/1000", "=T_REF", "=T_REF"),
            ("10월 생산계획", "=N_NOW", "=PL_B", "=PL_KG", "=E{r}*12/1000", "=CYC_ASIS", "=T_PREP+T_REF"),
            ("현재 Capa. 26.2톤", "=N_NOW", "=E{r}/KG_B", "=CAPA_NOW*1000/12", "=CAPA_NOW", "=8760/(CAPA_NOW*1000/KG_B)", "=T_PREP+T_REF"),
            ("기존 생산계획 ('27.5~)", "=N_NEW", "=E{r}/KG_B", f"={plan_b(2027, 5)}", "-", "=N_NEW*730/D{r}", "=T_PREP+T_REF"),
            ("Capa. 47.2톤 기준", "=N_NEW", "=E{r}/KG_B", "=CAPA_M", "=CAPA_RF", "=N_NEW*730/D{r}", "=T_PREP+T_REF"),
            ("2028 105℃ 50.2톤 기준", "=N_NEW", "=E{r}/KG_B", "=CAPA_M105", "=CAPA_105", "=N_NEW*730/D{r}", "=T_PREP+T_REF"))
for i, vals in enumerate(cmp_rows):
    rr = r + i
    for j, v in enumerate(vals):
        v = v.replace("{r}", str(rr)) if isinstance(v, str) else v
        put(ws, rr, 2 + j, v, "lab" if j == 0 else ("key" if j in (3, 5) else "calc"), [None, "int", "b", "kg", "t", "h", "h"][j])
    put(ws, rr, 9, f"=G{rr}-H{rr}", "key", "h")
REF["cmp0"] = (ws.title, f"B{CMP0}")
r += len(cmp_rows) + 1
section(ws, r, "⑤ 정제기 2대 운영 예시 — 완료 = 투입 + 5 Gal 실제 소요 (대기 포함)"); r += 1
put(ws, r, 2, "설비별 투입 간격 (월 Capa. 역산)", "lab")
put(ws, r, 3, '=IF(ISNUMBER(INT_NEXT),INT_NEXT,N_NEW*(8760/12)/(CAPA_M/KG_B))', "key", "h"); nm("INT_RF", ws, r, 3)
put(ws, r, 4, "2대 × 평균 월 730 h ÷ (월 Capa. ÷ 190)", "note"); r += 1
put(ws, r, 2, "설비별 대기 (주기 − 순수 47 h)", "lab"); put(ws, r, 3, "=INT_RF-T_PREP-T_REF", "key", "h"); nm("WAIT_RF", ws, r, 3); r += 1
put(ws, r, 2, "정제기 2 투입 시차 (가정)", "lab"); put(ws, r, 3, "=OFFSET_H", "calc", "h"); r += 1
put(ws, r, 2, "예시 월 · 시간", "lab"); put(ws, r, 3, "=TOBE_M", "calc", "d"); put(ws, r, 4, "=DAY(EOMONTH(TOBE_M,0))*24", "calc", "h0")
nm("TOBE_H", ws, r, 4); r += 1
for j, lab in enumerate(("순번", "정제기 1 투입 (h)", "완료 (h)", "완료일", "이 달 완료", "정제기 2 투입 (h)", "완료 (h)", "완료일", "이 달 완료")):
    put(ws, r, 2 + j, lab, "hdr")
r += 1
TB0 = r
for i in range(16):
    rr = r + i; k = i - 3
    put(ws, rr, 2, k, "calc", "int")
    for c0, off in ((3, "0"), (7, "OFFSET_H")):
        st, cp, cd, cf = (CL(c0 + x) for x in range(4))
        put(ws, rr, c0, f"={off}+B{rr}*INT_RF", "calc", "h")
        put(ws, rr, c0 + 1, f"={st}{rr}+LT_5G", "calc", "h")
        put(ws, rr, c0 + 2, f"=TOBE_M+INT({cp}{rr}/24)", "calc", "d")
        put(ws, rr, c0 + 3, f"=IF(AND({cp}{rr}>=0,{cp}{rr}<TOBE_H),1,0)", "calc", "int")
TB1 = r + 15
r = TB1 + 1
put(ws, r, 2, "이 달 완료 Batch / kg", "sub"); put(ws, r, 3, f"=SUM(F{TB0}:F{TB1})+SUM(J{TB0}:J{TB1})", "key", "int"); put(ws, r, 4, f"=C{r}*KG_B", "key", "kg")
REF["tobe_b"] = (ws.title, f"C{r}"); REF["tobe_kg"] = (ws.title, f"D{r}")
REF["tb0"] = (ws.title, f"B{TB0}"); REF["tobe_int"] = REF["INT_RF"]; REF["tobe_off"] = REF["OFFSET_H"]
r += 1
put(ws, r, 2, "월 Batch 상당 (월 Capa. ÷ 190)", "lab"); put(ws, r, 3, "=CAPA_M/KG_B", "key", "b"); REF["tobe_bm"] = (ws.title, f"C{r}")

# ---- 04_Capa기준
ws = wsK
title(ws, "04_Capa기준 | 연간 Capa. → 월 기준 · 운영 일정 · 2026 출하", "대정비 월이 있으므로 개선 후 월 Capa.는 (연간 − 대정비 월 생산) ÷ 11.")
widths(ws, {"B": 30, **{CL(i): 11 for i in range(3, 16)}})
section(ws, 4, "① Capa. 기준")
for j, lab in enumerate(("구분", "연간 (t)", "월 기준 (kg)", "Batch/월", "적용", "산식")):
    put(ws, 5, 2 + j, lab, "hdr", align="left" if j in (0, 5) else None)
caps = (("현재 (정제기 1대 · 103℃)", "=CAPA_NOW", "=CAPA_NOW*1000/12", "'27.1~3", "26.2 ÷ 12"),
        ("리플럭스 후 (정제기 2대)", "=CAPA_RF", "=CAPA_M", "'27.5~", "47,200 ÷ 11 · 대정비 월 1,188 kg 별도"),
        ("2028 105℃ 적용", "=CAPA_105", "=CAPA_M105", "'28.1~", "50,200 ÷ 11 · 대정비 월 1,188 kg 별도"),
        ("참고: 현재의 단순 2배", "=CAPA_NOW*2", "=C{r}*1000/12", "적용 안 함", "Mix·Premix 준비 등으로 2배 아님"))
for i, (lab, an, mo, when, basis) in enumerate(caps):
    rr = 6 + i
    put(ws, rr, 2, lab, "lab"); put(ws, rr, 3, an, "key", "t"); put(ws, rr, 4, mo.replace("{r}", str(rr)), "key", "kg1")
    put(ws, rr, 5, f"=D{rr}/KG_B", "calc", "b"); put(ws, rr, 6, when); put(ws, rr, 7, basis, "note")
put(ws, 10, 2, "대정비 월 생산 (매년 4월)", "lab"); put(ws, 10, 3, "=MAINT_KG/1000", "calc", "t"); put(ws, 10, 4, "=MAINT_KG", "key", "kg")
section(ws, 12, "② 운영 일정")
for j, lab in enumerate(("구분", "시작", "종료", "내용")):
    put(ws, 13, 2 + j, lab, "hdr", align="left" if j in (0, 3) else None)
sched = (("현재 정제기 1대", dt.date(2027, 1, 1), "=DATE(2027,MAINT_M,1)-1", "확정 생산 2,178 · 2,178 · 3,168 kg"),
         ("대정비 (매년)", "=DATE(2027,MAINT_M,1)", "=EOMONTH(DATE(2027,MAINT_M,1),0)", "생산 1,188 kg"),
         ("정제기 2대 생산", "=RF_START", "", "투입 시점을 엇갈려 병행 운전"),
         ("최초 12개월", "=RF_START", "=EDATE(RF_START,12)-1", "47.2톤 기준 (대정비 1개월 포함)"),
         ("ARS 200 L 충진", "=MIN(ARS_EZ_S,ARS_HS_S)", "", "이지켐·한솔 200 L · 5 Gal은 글로브 박스 유지"),
         ("105℃ (50.2톤)", "=T_105", "", "품질 검증·승인 필요"))
for i, (lab, s_, e_, t_) in enumerate(sched):
    rr = 14 + i
    put(ws, rr, 2, lab, "lab"); put(ws, rr, 3, s_, "calc", "d"); put(ws, rr, 4, e_, "calc", "d"); put(ws, rr, 5, t_, "note")
section(ws, 21, "③ 2026 출하 (제시분) vs 현재 Capa.")
month_hdr(ws, 22, 2026)
put(ws, 23, 2, "출하 (제시분)", "lab")
for m in MONTHS:
    put(ws, 23, 2 + m, f"={Iq}{CL(2 + m)}{SHIP[(2026, 'tot')]}", "calc", "kg")
put(ws, 23, 15, "=SUM(C23:N23)", "key", "kg")
put(ws, 24, 2, "현재 Capa. 월 환산", "lab")
for m in MONTHS:
    put(ws, 24, 2 + m, "=CAPA_NOW*1000/12", "calc", "kg")
put(ws, 24, 15, "=SUM(C24:N24)", "key", "kg")
put(ws, 25, 2, "Capa. − 출하", "lab")
for m in range(13):
    put(ws, 25, 3 + m, f"={CL(3 + m)}24-{CL(3 + m)}23", "calc", "sg")
put(ws, 26, 2, "CXMT·한솔 2026 물량 미제시 (합계 제외)", "note")
REF["sh26"] = (ws.title, "O23")

# =============================================================== 08_후공정부하 (05·06이 참조하므로 먼저 생성)
wsA = wb.create_sheet("05_월별_Capa기준")
wsB = wb.create_sheet("06_월별_생산계획")
wsV = wb.create_sheet("07_버전비교")
wsL = wb.create_sheet("08_후공정부하")
ws = wsL
title(ws, "08_후공정부하 | 5 Gal · 200 L 충진 · OQC · 검사 (출하계획 기준)", "생산 버전과 무관하게 출하량으로 계산. ARS 시간 미입력 시 수동 8 h 적용.")
widths(ws, {"B": 30, **{CL(i): 9.5 for i in range(3, 16)}})
LOAD = {}
r = 4
load_rows = (("g5kg", "5 Gal 출하 (하이닉스+CXMT) kg", "=N({hx})+N({cx})", "kg"),
             ("g5bt", "5 Gal 병", "={g5kg}/KG_BOT", "b"),
             ("g5h", "5 Gal 충진 h", "={g5bt}*FILL_5G", "h"),
             ("g5oqc", "5 Gal OQC h", "={g5bt}/BOT_B*OQC_5G", "h"),
             ("gbav", "글로브 박스 가용 h (1교대)", "=WH_DAY*WD_MON", "h0"),
             ("gbld", "글로브 박스 부하율", "={g5h}/{gbav}", "pct"),
             ("ezc", "이지켐 용기", "=N({ez})/KG_EZ", "b"),
             ("ezm", "이지켐 충진 방식", '=IF({ezc}=0,"-",IF({d}>=ARS_EZ_S,"ARS","수동"))', None),
             ("ezh", "이지켐 충진 h", '={ezc}*IF(AND({ezm}="ARS",ISNUMBER(ARS_EZ_T)),ARS_EZ_T,FILL_EZ)', "h"),
             ("hsc", "한솔 용기", "=N({hs})/KG_HS", "b"),
             ("hsm", "한솔 충진 방식", '=IF({hsc}=0,"-",IF({d}>=ARS_HS_S,"ARS","수동"))', None),
             ("hsh", "한솔 충진 h", '={hsc}*IF(AND({hsm}="ARS",ISNUMBER(ARS_HS_T)),ARS_HS_T,FILL_HS)', "h"),
             ("l2h", "200 L 충진 h 합계", "={ezh}+{hsh}", "h"),
             ("l2oqc", "200 L OQC h", "=({ezc}+{hsc})*OQC_200", "h"),
             ("nb", "필요 Batch (출하 ÷ 190)", "=(N({hx})+N({cx})+N({ez})+N({hs}))/KG_B", "b"),
             ("qch", "PQC~이송·FQC h (필요 Batch 기준)", "={nb}*(T_PQC+T_FQC)", "h"))
for yr in YEARS:
    month_hdr(ws, r, yr, label=f"{yr}", total="연간"); hr = r; r += 1
    pos = {}
    for key, lab, fx, nf in load_rows:
        pos[key] = r; r += 1
    for key, lab, fx, nf in load_rows:
        rr = pos[key]
        put(ws, rr, 2, lab, "lab")
        for m in MONTHS:
            col = CL(2 + m)
            sub = {k2: f"{col}{v}" for k2, v in pos.items()}
            sub.update({c: ship(yr, c, m) for c, _ in CUST}); sub["d"] = f"{col}${hr}"
            put(ws, rr, 2 + m, fx.format(**sub), "key" if key in ("g5h", "l2h") else "calc", nf)
        tot = {"ezm": "-", "hsm": "-", "gbld": f"=MAX(C{rr}:N{rr})", "gbav": f"=C{rr}"}.get(key, f"=SUM(C{rr}:N{rr})")
        put(ws, rr, 15, tot, "key", nf)
    ws.conditional_formatting.add(f"C{pos['gbld']}:N{pos['gbld']}", CellIsRule(operator="greaterThan", formula=["1"], font=Font(color=RED, bold=True)))
    LOAD[yr] = pos; r += 1
put(ws, r, 2, "연간 열의 부하율은 월 최대", "note")
Lq = q(wsL)

# =============================================================== 05 · 06 월별 (두 버전)
MROW = {}


def month_sheet(ws, ver):
    title(ws, ("05_월별_Capa기준 | 2027 47.2톤 · 2028 50.2톤 기준 생산 vs 출하" if ver == "A"
               else "06_월별_생산계획 | 기존 생산계획 기준 생산 vs 출하"),
          "PPT 연도별 표. 생산 = " + ("1~4월 확정 + 월 Capa. (연간 ÷ 11)" if ver == "A" else "01_입력 G 기존 생산계획") + " · 출하·충진은 01_입력 · 08_후공정부하 연결.")
    widths(ws, {"B": 30, **{CL(i): 9.5 for i in range(3, 16)}})
    rr = 4
    pos_all = {}
    for yr in (2027, 2028):
        month_hdr(ws, rr, yr, label=f"{yr}"); hr = rr; rr += 1
        keys = ("cond", "plan", "pb", "diff", "inv", "hx", "cx", "ez", "hs", "ship", "nb", "l2m", "g5h", "l2h", "ezc", "hsc")
        pos = {k: rr + i for i, k in enumerate(keys)}
        labels = {"cond": "운전 조건", "plan": "생산 (kg)", "pb": "Batch 환산 (÷190)", "hx": "SK하이닉스", "cx": "CXMT", "ez": "이지켐",
                  "hs": "한솔", "ship": "출하 합계", "diff": "생산 − 출하", "inv": "재고 (전월 재고 + 생산 − 출하)", "nb": "필요 Batch (출하÷190)",
                  "l2m": "200 L 충진 방식 (이지켐/한솔)",
                  "g5h": "5 Gal 충진 h", "l2h": "200 L 충진 h", "ezc": "이지켐 용기", "hsc": "한솔 용기",
                  }
        for k in keys:
            put(ws, pos[k], 2, labels[k], "sub" if k in ("plan", "ship", "inv") else "lab")
        for m in MONTHS:
            c = CL(2 + m); d = f"{c}${hr}"
            if ver == "A":
                cond = (f'=IF(MONTH({d})=MAINT_M,"대정비",IF({d}<RF_START,"정제기 1대",IF({d}>=T_105,"2대·105℃","정제기 2대")))')
                plan = (f"={plan_b(yr, m)}" if (yr == 2027 and m <= 4) else
                        f"=IF(MONTH({d})=MAINT_M,MAINT_KG,IF({d}>=T_105,CAPA_M105,IF({d}>=RF_START,CAPA_M,CAPA_NOW*1000/12)))")
            else:
                cond = f'=IF(MONTH({d})=MAINT_M,"대정비",IF({d}<RF_START,"정제기 1대","정제기 2대"))'
                plan = f"={plan_b(yr, m)}"
            previ = (f"{CL(1 + m)}{pos['inv']}" if m > 1 else (f"N{pos_all[2027]['inv']}" if yr == 2028 else None))
            vals = {"cond": cond, "plan": plan, "pb": f"={c}{pos['plan']}/KG_B",
                    "ship": f"=SUM({c}{pos['hx']}:{c}{pos['hs']})", "diff": f"={c}{pos['plan']}-{c}{pos['ship']}",
                    "nb": f"={c}{pos['ship']}/KG_B",
                    "inv": f"={previ if previ else 'N(INV_0)'}+{c}{pos['plan']}*IF(ISNUMBER(YIELD_Q),YIELD_Q,1)-{c}{pos['ship']}",
                    "l2m": f'={Lq}{c}{LOAD[yr]["ezm"]}&"/"&{Lq}{c}{LOAD[yr]["hsm"]}'}
            for k_, cu in (("hx", "hx"), ("cx", "cx"), ("ez", "ez"), ("hs", "hs")):
                vals[k_] = f"={ship(yr, cu, m)}"
            for k_ in ("g5h", "l2h", "ezc", "hsc"):
                vals[k_] = f"={Lq}{c}{LOAD[yr][k_]}"
            for k_, v in vals.items():
                nf = {"plan": "kg1", "pb": "b", "hx": "kg1", "ship": "kg1", "diff": "sg", "inv": "sg", "nb": "b", "g5h": "h", "l2h": "h",
                      "ezc": "b", "hsc": "b"}.get(k_, "kg")
                if k_ in ("cond", "l2m"):
                    nf = None
                put(ws, pos[k_], 2 + m, v, "key" if k_ in ("plan", "inv") else "calc", nf)
        for k in keys:
            nf = {"pb": "b", "nb": "b", "g5h": "h", "l2h": "h", "ezc": "b", "hsc": "b", "diff": "sg", "inv": "sg"}.get(k, "kg")
            tot = {"cond": "", "l2m": "-", "inv": f"=N{pos['inv']}"}.get(k, f"=SUM(C{pos[k]}:N{pos[k]})")
            put(ws, pos[k], 15, tot, "key", nf)
        ws.conditional_formatting.add(f"C{pos['cond']}:N{pos['cond']}", FormulaRule(formula=[f'C{pos["cond"]}="대정비"'], fill=fill(F_SUB), font=Font(color=RED, bold=True)))
        ws.conditional_formatting.add(f"C{pos['plan']}:N{pos['plan']}", FormulaRule(formula=[f'C{pos["cond"]}="대정비"'], fill=fill(F_SUB)))
        if yr == 2027:
            for m in range(1, 5):
                ws.cell(row=pos["plan"], column=2 + m).fill = fill(F_IN)
        pos["hdr"] = hr
        pos_all[yr] = pos
        rr = pos["hsc"] + 2
    put(ws, rr, 2, ("2027년 1~4월은 확정 생산 (01_입력 G), 5월부터 47,200 ÷ 11 · 2028년 50,200 ÷ 11 · 4월 대정비 1,188 kg · 재고는 기초재고 2,882 kg부터" if ver == "A"
                    else "01_입력 G 기존 생산계획 · 2027년 1~4월 확정 · 2028년 생산계획 38,412 kg"), "note")
    return pos_all


MROW["A"] = month_sheet(wsA, "A")
MROW["B"] = month_sheet(wsB, "B")

# =============================================================== 07_버전비교
ws = wsV
title(ws, "07_버전비교 | Capa. 기준 vs 기존 생산계획 (2027~2028)", "재고 = 기초재고 (2027-01-01) + 생산 − 출하 · 2028년은 2027년 말 재고에서 이어짐.")
widths(ws, {"B": 30, **{CL(i): 10 for i in range(3, 27)}})
section(ws, 4, "① 연간 비교 (kg)")
hdrs = ("버전", "2027 생산", "2027 출하", "2027 차이", "2028 생산", "2028 출하", "2028 차이", "'27 말 재고", "'28 말 재고", "최저 재고", "최저 시점")
for j, lab in enumerate(hdrs):
    put(ws, 5, 2 + j, lab, "hdr", align="left" if j == 0 else None)
for i, (ver, lab, wsx) in enumerate((("A", "① Capa. 기준 (47.2 / 50.2톤)", wsA), ("B", "② 기존 생산계획", wsB))):
    rr = 6 + i; P = MROW[ver]; X = q(wsx)
    c27, c28 = P[2027]["inv"], P[2028]["inv"]
    rng27, rng28 = f"{X}$C${c27}:$N${c27}", f"{X}$C${c28}:$N${c28}"
    h27, h28 = f"{X}$C${P[2027]['hdr']}:$N${P[2027]['hdr']}", f"{X}$C${P[2028]['hdr']}:$N${P[2028]['hdr']}"
    vals = [lab, f"={X}O{P[2027]['plan']}", f"={X}O{P[2027]['ship']}", f"=C{rr}-D{rr}", f"={X}O{P[2028]['plan']}", f"={X}O{P[2028]['ship']}",
            f"=F{rr}-G{rr}", f"={X}N{c27}", f"={X}N{c28}", f"=MIN({rng27},{rng28})",
            f"=IF(MIN({rng27})<=MIN({rng28}),INDEX({h27},MATCH(MIN({rng27}),{rng27},0)),INDEX({h28},MATCH(MIN({rng28}),{rng28},0)))"]
    for j, v in enumerate(vals):
        put(ws, rr, 2 + j, v, "lab" if j == 0 else ("key" if j in (3, 6, 7, 8, 9) else "calc"),
            [None, "kg", "kg", "sg", "kg", "kg", "sg", "sg", "sg", "sg", "ym"][j])
    REF[f"v{ver}"] = (ws.title, f"B{rr}")
put(ws, 8, 2, "기초재고 (2027-01-01)", "lab"); put(ws, 8, 3, "=INV_0", "calc", "kg")
section(ws, 10, "② 월별 생산 · 출하 · 재고 (kg)")
put(ws, 11, 2, "월", "hdr", align="left")
for k in range(24):
    yr, m = 2027 + k // 12, k % 12 + 1
    put(ws, 11, 3 + k, dt.date(yr, m, 1), "hdr", "ym")
series = (("출하", "A", "ship"), ("① 생산 (Capa. 기준)", "A", "plan"), ("① 재고", "A", "inv"),
          ("② 생산 (기존 생산계획)", "B", "plan"), ("② 재고", "B", "inv"))
for i, (lab, ver, key) in enumerate(series):
    rr = 12 + i; wsx = wsA if ver == "A" else wsB
    put(ws, rr, 2, lab, "sub" if key == "inv" else "lab")
    for k in range(24):
        yr, m = 2027 + k // 12, k % 12 + 1
        put(ws, rr, 3 + k, f"={q(wsx)}{CL(2 + m)}{MROW[ver][yr][key]}", "key" if key == "inv" else "calc", "sg" if key == "inv" else "kg")
REF["series"] = (ws.title, "B12")

# =============================================================== 09_확인사항
ws = wb.create_sheet("09_확인사항")
title(ws, "09_확인사항 | 생산팀 확인 내용 · 조건별 생산량 입력", "생산·품질·설비·충진 관련 내용만 정리.")
widths(ws, {"B": 22, "C": 52, "D": 40, "E": 34})
section(ws, 4, "① 생산팀 확인 내용")
for j, lab in enumerate(("구분", "생산팀 설명", "자료 반영", "확인 필요")):
    put(ws, 5, 2 + j, lab, "hdr", align="left")
conv = (("공정시간", "57 h에서 IQC·준비·PQC~이송·FQC를 빼면 순수 정제 45 h · 대화 중 57 − 2 − 6 = 49 h 계산도 있음", "IQC~FQC 57 h, 순수 정제 45 h (차감 계산)", "원자료 시간 측정의 시작점·종료점"),
        ("과거 시간", "과거 약 72 h, 최근 약 53 h 언급", "원자료 설명으로 보존 (57 h 기준 유지)", "두 값의 측정 범위가 같은지"),
        ("정제기 2대", "정제기 1 투입 후 초류 진행에 맞춰 정제기 2 투입 (다음 날 / 초류 중간 / 초류 종료 무렵 언급)", "투입 시점을 엇갈려 병행 운전 · 운전 예시 시차 30 h는 가정", "정제기 2 투입 시점 · 설비 번호(2070·2060) 도면 대조 · 개조/신규 여부"),
        ("생산물 도착 간격", "57 ÷ 2 = 28.5 h는 이상적인 경우의 도착 간격 설명", "정제기별 Batch 시간으로 사용하지 않음", "-"),
        ("Capa. 2배 아님", "추가 투입용 Mix·Premix 준비 등에 설비 시간이 필요", "47.2톤 (현재 26.2톤의 2배 52.4톤 아님)", "Mix·Premix 준비시간 · 사용 설비"),
        ("리플럭스 목적", "하이닉스 요구 품질 대응을 위한 정제 기능 개선 · 환류 운전은 회수 속도와 Batch 시간에 영향", "품질 대응 + 생산능력 확대 · 온도 상승 효과와 2대 효과 구분", "환류 운전 시 Batch 시간"),
        ("회수량", "조건별 약 160~170 kg / 추가 투입 시 약 220 kg / 전체 약 195~200 kg (197 kg) / Crude 200 kg × 75~80% ≈ 150~160 kg", "환산 기준 190 kg/Batch 유지 · 조건별 실측 입력란", "조건별 투입·회수·양품·잔량 실측"),
        ("수율 표현", "\"65%에서 5% 올렸다\"", "65→70% 개선으로 단정하지 않음", "수치 정의"),
        ("135 · 120", "Mix 유무에 따라 언급", "kg 생산량으로 입력하지 않음", "단위 · 색도 등 품질 지표인지"),
        ("이지켐 색도", "색도 조건 때문에 추가 투입 제한", "고객별 품질 조건 별도", "고객별 품질·색도 규격 · 합격률"),
        ("5 Gal 충진", "9병 18 h 작업, 근무시간 반영 시 약 2일 · 개선 방향 1~1.5일 · 18 h → 9 h 방안 검토", "2 h/병 · 18 h 유지 (개선은 미반영)", "교대·연장근무 · 충진 인원"),
        ("ARS", "ARS 자체 충진은 더 길 수 있으나 용기 투입·반출·퍼지 포함 시 전체 시간은 비슷할 수 있음", "ARS 시간 미입력 시 수동 8 h 적용 · Capa. 가산 없음", "ARS 작업 단계별 시간 (용기 준비~설비 전환)"),
        ("한솔 충진", "200 L 수동 충진", "8 h/용기 반영", "ARS 전환 후 시간"),
        ("출하 물량", "대화 중 '약 40톤' 언급", "고객별 출하계획 (2027 41.72 / 2028 45.76톤) 유지", "-"))
for i, row_ in enumerate(conv):
    rr = 6 + i
    for j, v in enumerate(row_):
        put(ws, rr, 2 + j, v, "lab" if j == 0 else "calc", align="left", wrap=True)
    ws.row_dimensions[rr].height = 32
r = 6 + len(conv) + 1
section(ws, r, "② 조건별 생산량 (실측 입력 — 구두 언급치는 참고)"); r += 1
for j, lab in enumerate(("조건", "구두 언급", "신규 Crude (kg)", "재투입·Mix (kg)", "총 투입 (kg)", "회수량 (kg)", "양품량 (kg)", "잔량·이월 (kg)", "회수율 (회수÷총투입)")):
    put(ws, r, 2 + j, lab, "hdr", align="left" if j < 2 else None)
widths(ws, {"F": 12, "G": 12, "H": 12, "I": 12, "J": 16})
r += 1
for lab, oral in (("일부 운전 조건", "약 160~170 kg"), ("추가 투입 · Make-up", "약 220 kg"), ("전체 평균", "약 195~200 kg (197 kg)"),
                  ("Crude만 (참고)", "200 kg × 75~80% ≈ 150~160 kg"), ("이지켐 (색도 제한)", "추가 투입 제한")):
    put(ws, r, 2, lab, "lab"); put(ws, r, 3, oral, "calc", align="left")
    for c in (4, 5, 7, 8, 9):
        put(ws, r, c, None, "ask", "kg")
    put(ws, r, 6, f'=IF(COUNT(D{r}:E{r})=0,"",N(D{r})+N(E{r}))', "calc", "kg")
    put(ws, r, 10, f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(G{r})),G{r}/F{r},"")', "calc", "pct")
    r += 1
put(ws, r, 2, "수율은 총 투입(신규 + 재투입) 기준 · 신규 Crude만 분모로 쓰지 않음", "note")

# =============================================================== 10_과거Batch_참고
ws = wb.create_sheet("10_과거Batch_참고")
title(ws, "10_과거Batch_참고 | 2026 생산계획 표시일자 (원자료 보존 · 참고용)", "표시일자 = 해당 Batch의 생산 및 충진 완료 예정일 (생산계획 · 실적 아님).")
widths(ws, {"B": 8, "C": 12, "D": 8, "F": 10, "G": 12, "H": 12})
L = [(37, 5, 23), (38, 5, 26), (39, 6, 2), (40, 6, 5), (41, 6, 8), (42, 6, 11), (43, 6, 14), (44, 6, 17), (45, 6, 20), (46, 6, 23),
     (47, 6, 25), (48, 6, 27), (49, 6, 28), (50, 7, 1), (51, 7, 3), (52, 7, 6), (53, 7, 8), (54, 7, 11), (55, 7, 14), (56, 7, 17),
     (57, 7, 20), (58, 7, 23), (59, 7, 26), (60, 7, 29), (61, 8, 1), (62, 8, 4), (63, 8, 7), (64, 8, 10), (65, 8, 13), (66, 8, 16),
     (67, 8, 19), (68, 8, 22), (69, 8, 25), (70, 8, 28), (71, 8, 31), (72, 9, 3), (73, 9, 5), (74, 9, 7), (75, 9, 9), (76, 9, 11),
     (77, 9, 14), (78, 9, 19), (79, 9, 23), (80, 9, 26), (81, 9, 28), (82, 10, 1), (83, 10, 4), (84, 10, 7), (85, 10, 10),
     (86, 10, 12), (87, 10, 15), (88, 10, 18), (89, 10, 20), (90, 10, 23), (91, 10, 26), (92, 10, 29)]
for j, lab in enumerate(("Batch", "표시일자", "월")):
    put(ws, 4, 2 + j, lab, "hdr")
for i, (n, m, d) in enumerate(L):
    rr = 5 + i
    put(ws, rr, 2, n, "in", "int"); put(ws, rr, 3, dt.date(2026, m, d), "in", "d"); put(ws, rr, 4, f"=MONTH(C{rr})", "calc", "int")
last = 4 + len(L)
for j, lab in enumerate(("월", "번호 개수", "원자료 Total")):
    put(ws, 4, 6 + j, lab, "hdr")
for i, (m, tot) in enumerate(((5, None), (6, 8), (7, 11), (8, 8), (9, 8), (10, 11))):
    rr = 5 + i
    put(ws, rr, 6, f"{m}월", "lab"); put(ws, rr, 7, f"=COUNTIF($D$5:$D${last},{m})", "calc", "int")
    put(ws, rr, 8, tot, "in" if tot else "ask", "int")
put(ws, 12, 6, "5월은 #37·#38 일부만 제공", "note")

# =============================================================== 마무리
order = ["01_입력", "02_공정시간", "03_Batch주기", "04_Capa기준", "05_월별_Capa기준", "06_월별_생산계획", "07_버전비교", "08_후공정부하",
         "09_확인사항", "10_과거Batch_참고"]
wb._sheets = [wb[n] for n in order]
tabs = {"01_입력": "2E75B6", "02_공정시간": "EB002C", "03_Batch주기": "EB002C", "04_Capa기준": "EB002C", "05_월별_Capa기준": "FF7900",
        "06_월별_생산계획": "FF7900", "07_버전비교": "FF7900", "08_후공정부하": "70AD47", "09_확인사항": "7F7F7F", "10_과거Batch_참고": "BFBFBF"}
for w in wb.worksheets:
    w.sheet_properties.tabColor = tabs[w.title]
    w.sheet_view.zoomScale = 90
    if w.title[:2] in ("05", "06", "07", "08"):
        w.freeze_panes = "C4"
wb.active = 0
wb.save(OUT)
json.dump({"REF": REF, "MROW": {v: {str(y): p for y, p in d.items()} for v, d in MROW.items()},
           "LOAD": {str(y): p for y, p in LOAD.items()}, "PLAN": PLAN, "SHIP": {f"{k[0]}_{k[1]}": v for k, v in SHIP.items()}},
          open(OUT + ".map.json", "w"), ensure_ascii=False)
print("saved", OUT)
