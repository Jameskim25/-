# -*- coding: utf-8 -*-
"""SKTC CpZr 생산능력 계산 파일 생성 스크립트.

원자료(첨부 PPT 3장 발표자 노트 + 분포도)와 사용자 제공 조건을 입력값으로 두고,
모든 결과는 Excel 수식으로 계산한다. 결과 값은 LibreOffice 재계산 후 확인한다.
"""
import datetime as dt
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.formula import ArrayFormula
from pptx import Presentation

SRC_PPTX = sys.argv[1]
OUT = sys.argv[2]
SRC_NAME = "SKTC_CpZr_Capacity_Reflux_Roadmap_2027_2028.pptx (첨부 원본)"
USER_SRC = "사용자 제공 조건 (2026-09-30 지시)"

# ---------------------------------------------------------------- styles
FONT = "Arial"
def font(color="000000", bold=False, italic=False, size=9):
    return Font(name=FONT, color=color, bold=bold, italic=italic, size=size)

FILL = {
    "in": PatternFill("solid", fgColor="DDEBF7"),     # 제공/확정 입력
    "unk": PatternFill("solid", fgColor="FFF2CC"),    # 미확인 입력(공란)
    "est": PatternFill("solid", fgColor="FCE4D6"),    # 추정·시나리오 가정 입력
    "rev": PatternFill("solid", fgColor="E4DFEC"),    # 검토안
    "hdr": PatternFill("solid", fgColor="404040"),
    "hdr2": PatternFill("solid", fgColor="C55A11"),   # 분석/추정 열 머리글
    "sec": PatternFill("solid", fgColor="F2F2F2"),
    "key": PatternFill("solid", fgColor="FDE9E7"),    # 핵심 결과
    "none": PatternFill(fill_type=None),
}
BLUE, GREEN, BLACK, GRAY = "0000FF", "008000", "000000", "7F7F7F"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

NF = {"kg": '#,##0.0;[Red]-#,##0.0;"0"', "kg0": '#,##0;[Red]-#,##0;"0"', "h": '0.0;[Red]-0.0', "h2": '0.00',
      "b": '0.00', "b0": '0', "t": '0.00', "pct": '0.0%', "d": 'yyyy-mm-dd', "ym": 'yyyy-mm', "md": 'm"/"d', "int": '0'}

wb = Workbook()

def style_cell(c, kind="calc", nf=None, bold=False):
    c.border = BORDER
    c.alignment = WRAP
    if kind == "in":
        c.font = font(BLUE, bold); c.fill = FILL["in"]
    elif kind == "unk":
        c.font = font(BLUE, bold); c.fill = FILL["unk"]
    elif kind == "est":
        c.font = font(BLUE, bold); c.fill = FILL["est"]
    elif kind == "rev":
        c.font = font(BLUE, bold); c.fill = FILL["rev"]
    elif kind == "link":
        c.font = font(GREEN, bold)
    elif kind == "key":
        c.font = font(BLACK, True); c.fill = FILL["key"]
    elif kind == "text":
        c.font = font(BLACK, bold)
    elif kind == "note":
        c.font = font(GRAY, italic=True)
    else:
        c.font = font(BLACK, bold)
    if nf:
        c.number_format = NF.get(nf, nf)

def header(ws, row, labels, col=1, fill="hdr", height=30):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font = font("FFFFFF", True); c.fill = FILL[fill]; c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[row].height = height

def title(ws, text, sub):
    ws["A1"] = text; ws["A1"].font = font("EB002C", True, size=14)
    ws["A2"] = sub; ws["A2"].font = font(GRAY, italic=True, size=9)

def legend(ws, row, col=1):
    items = [("in", "입력값 (제공·확정·계획) — 파란 글자"), ("unk", "미확인 입력란 (공란 유지, 0 치환 금지)"),
             ("est", "추정·시나리오 가정 입력"), ("rev", "검토안 (미확정)"), ("calc", "수식 (검정) / 다른 시트 참조 (초록)"),
             ("key", "핵심 결과 (수식)")]
    for i, (k, lab) in enumerate(items):
        c = ws.cell(row=row, column=col + 2 * i, value="  ")
        style_cell(c, k if k != "calc" else "calc")
        d = ws.cell(row=row, column=col + 2 * i + 1, value=lab); d.font = font(GRAY, size=8)

def widths(ws, spec):
    for k, v in spec.items():
        ws.column_dimensions[k].width = v

# ======================================================== 01_Inputs
wsI = wb.active; wsI.title = "01_Inputs"
title(wsI, "01_Inputs | CpZr 생산능력 계산 입력값", "입력값을 바꾸면 03~06 시트가 자동 갱신됩니다. 미확인 항목은 공란(노란 칸)으로 두고 확인 후 입력하세요.")
legend(wsI, 3)
header(wsI, 5, ["항목", "값", "단위", "적용 범위", "구분", "출처", "비고"])
REF = {}   # key -> absolute ref "'01_Inputs'!$B$n"
ROWS_I = []
r = 6

def sec(label):
    global r
    for col in range(1, 8):
        c = wsI.cell(row=r, column=col, value=label if col == 1 else None)
        c.fill = FILL["sec"]; c.font = font(BLACK, True); c.border = BORDER
    r += 1

def inp(key, item, val, unit, scope, kind, src, note="", nf=None):
    """kind: 확정/제공/계획/추정/미확인/검토안/수식"""
    global r
    wsI.cell(row=r, column=1, value=item)
    c = wsI.cell(row=r, column=2, value=val)
    wsI.cell(row=r, column=3, value=unit)
    wsI.cell(row=r, column=4, value=scope)
    wsI.cell(row=r, column=5, value=kind)
    wsI.cell(row=r, column=6, value=src)
    wsI.cell(row=r, column=7, value=note)
    for col in (1, 3, 4, 5, 6, 7):
        style_cell(wsI.cell(row=r, column=col), "text")
    k = {"확정": "in", "제공": "in", "계획": "in", "추정": "est", "미확인": "unk", "검토안": "rev", "수식": "calc"}[kind]
    style_cell(c, k, nf)
    if isinstance(val, (dt.date, dt.datetime)):
        c.number_format = NF["d"]
    REF[key] = f"'01_Inputs'!$B${r}"
    r += 1

sec("A. 설비 · 가동시간 (정제기)")
inp("n_ref", "CpZr 정제기 수", 1, "대", "하이닉스·CXMT·이지켐 공동 사용", "제공", USER_SRC, "리플럭스 적용 후에도 동일 1대 (증설 아님)", "int")
inp("temp_now", "현재 정제온도", 103, "℃", "현재 조건", "제공", USER_SRC, "", "int")
inp("mdays", "월 비교 기준 일수", 30, "일/월", "정제기 월 비교 기준", "제공", USER_SRC, "30일×24시간", "int")
inp("mh", "월 비교 기준 시간 (정제기)", f"=B{r-1}*24", "h/월", "정제기 전용 — 후공정 근무시간에 자동 적용하지 않음", "수식", "수식", "", "h")
inp("yh", "연간 달력시간 (365일)", "=365*24", "h/년", "연간 연속 운전 환산", "수식", "수식", "정지·보수 미반영 상한", "int")
inp("down", "정제기 계획 정지·보수시간", None, "h/월", "월별 입력은 04 시트 F열", "미확인", "-", "미입력 시 04 시트에서 '정지 미반영'으로 표시")

sec("B. 공정시간")
inp("t_iqc", "IQC (수입검사)", 2, "h/회", "원료 Lot당 (여러 Batch 공통 적용 여부 미확인)", "제공", USER_SRC, "NMR·ICP·IC", "h")
inp("n_iqc", "IQC 1회당 적용 Batch 수", None, "Batch/Lot", "원료 Lot ↔ Batch 관계", "미확인", "-", "미입력 시 Batch마다 1회(상한)로 작업량 계산")
inp("t_prep", "준비·투입 (Canister→정제기)", None, "h/Batch", "정제 57 h 포함 여부 확인", "미확인", "-", "정제시간에 포함되면 공란 유지(중복 가산 금지)")
inp("t_ref", "정제시간 (현재 103℃)", 57, "h/Batch", "정제기 점유 — 가열·냉각·배출 포함 범위 확인", "제공", USER_SRC, "포함된 시간은 다시 가산하지 않음", "h")
inp("t_pqc", "PQC (공정검사, NMR)", None, "h/Batch", "", "미확인", "-")
inp("t_tank", "Product Tank 이송", None, "h/Batch", "", "미확인", "-")
inp("t_fqc", "FQC (제품검사)", 2, "h/Batch", "검사시간 (대기시간 제외)", "제공", USER_SRC, "NMR·ICP-MS·IC·점도", "h")
inp("t_fqcw", "FQC 검사 대기시간", None, "h/Batch", "검사시간과 구분", "미확인", "-")
inp("t_clean", "세척·전환 (고객/제품 전환)", None, "h/회", "", "미확인", "-")
inp("t_5g", "5 Gal 충진 (글로브 박스)", 2, "h/병", "하이닉스·CXMT, 필터 포함", "제공", USER_SRC, "필터는 충진에 포함 — 별도 시간 가산 없음", "h")
inp("n_5g_b", "5 Gal Batch 충진 병수 (작업 기준)", 8, "병/Batch", "약 8병 — 평균 충진 수량 여부 확인", "제공", USER_SRC, "8병×20 kg = 160 kg (190 kg과 30 kg 차이 확인)", "int")
inp("t_5g_b", "5 Gal Batch 충진시간 (작업 기준)", 16, "h/Batch", "약 8병 기준 — 190 kg 전량 충진시간과 동일시하지 않음", "제공", USER_SRC, "", "h")
inp("chk_5g", "5 Gal 정합 확인 (병수×병당시간)", f"=B{r-2}*B{r-3}", "h/Batch", "16 h와 일치 여부", "수식", "수식", "", "h")
inp("t_200", "200 L 충진 (수동, 필터 포함)", 8, "h", "용기당/Batch당 적용 범위 미확인", "제공", USER_SRC, "이지켐, 140 kg/용기", "h")
inp("sw_200", "200 L 8 h 적용 단위 (1=용기당, 2=Batch당)", 1, "선택", "04 시트 200 L 작업시간 계산", "추정", "보수적 가정", "확인 후 1 또는 2로 수정", "int")
inp("t_ars", "200 L ARS 자동 충진시간", None, "h/용기", "2027년~ 이지켐", "미확인", "-", "미입력 시 수동 8 h를 참고값으로 사용·표시")
inp("t_oqc", "OQC·출하 작업", 2, "h/회", "출하 Lot당 (적용 단위 미확인) — 물류 운송시간 제외", "제공", USER_SRC, "", "h")

sec("C. 생산량 · 용기")
inp("kg_b", "Batch당 생산량 (환산 기준)", 190, "kg/Batch", "생산량 환산 기준 — 정제 회수량/최종 양품량 여부 확인", "제공", USER_SRC, "", "kg0")
inp("kg_5g", "5 Gal 충진량", 20, "kg/병", "하이닉스·CXMT (5 Gal, 약 19 L)", "제공", USER_SRC, "", "kg0")
inp("L_5g", "5 Gal 용기 용량", 19, "L", "", "제공", USER_SRC, "", "int")
inp("kg_200", "200 L 충진량", 140, "kg/용기", "이지켐", "제공", USER_SRC, "", "kg0")
inp("kg_b_rf", "리플럭스 후 Batch당 생산량", None, "kg/Batch", "2027.7~", "미확인", "-", "미입력 시 190 kg 유지로 계산하고 04 시트에 표시")

sec("D. Capa. (연간 환산 기준값)")
inp("capa_now", "현재 연간 Capa. (103℃)", 26.2, "t/년", "현재 조건 기준값 — 유지", "제공", USER_SRC, "시간·가동률을 조정해 맞추지 않음", "t")
inp("capa_add", "리플럭스 증가분", 21, "t/년", "'27.7~'28.6 최초 12개월 기준", "제공", USER_SRC, "", "t")
inp("capa_rf", "리플럭스 후 연간 환산 Capa.", f"=B{r-2}+B{r-1}", "t/년", "연간 환산값 ≠ 2027년 실제 생산량", "수식", "수식", "ARS 효과 미가산", "t")
inp("capa_tadd", "105℃ 온도 상승 예상 증가분", 3, "t/년", "2028 검토안 — 미확정·검증·승인 필요", "검토안", USER_SRC, "적용 시점 미정", "t")
inp("capa_t", "105℃ 적용 시 연간 환산 Capa.", f"=B{r-2}+B{r-1}", "t/년", "검토안 (산술 계산값)", "수식", "수식", "ARS 효과 미가산", "t")
inp("target", "목표 Capa. (2028년 이후)", 50, "t/년", "50톤 수준", "계획", USER_SRC, "", "t")
inp("max11", "과거 월 최대 생산 Batch", 11, "Batch/월", "히스토리 최대 기준", "제공", USER_SRC, "", "int")
inp("t_ref_rf", "리플럭스 후 정제시간", None, "h/Batch", "2027.7~", "미확인", "-", "현재 57 h로 47.2톤을 계산하지 않음")
inp("ovh_rf", "리플럭스 후 비정제 시간 (간격−정제)", None, "h/Batch", "2027.7~", "미확인", "-", "미입력 시 현재 관측 추정값(03 시트) 적용 — 04 시트에 표시")
inp("temp_new", "검토 정제온도", 105, "℃", "2028 검토안", "검토안", USER_SRC, "열 안정성·Dimer·Unknown impurity·수율 검증 필요", "int")
inp("temp_date", "105℃ 적용 시점", None, "날짜", "", "미확인", "-", "적용 시점 미정")

sec("E. 일정 (리플럭스 · ARS)")
inp("ars_inst", "ARS 설치 완료 (2026년 내)", dt.date(2026, 12, 31), "날짜", "200 L 이지켐", "계획", USER_SRC, "설치일 미확정 — 2026년 내")
inp("ars_start", "ARS 운영 시작", dt.date(2027, 1, 1), "날짜", "200 L 이지켐", "계획", USER_SRC, "운영 개시일 확인 필요")
inp("con_s", "리플럭스 공사 시작", dt.date(2027, 2, 1), "날짜", "", "계획", USER_SRC, "2027년 2~5월")
inp("con_e", "리플럭스 공사 종료", dt.date(2027, 5, 31), "날짜", "", "계획", USER_SRC)
inp("trial_s", "시운전 시작", dt.date(2027, 6, 1), "날짜", "", "계획", USER_SRC, "2027년 6월")
inp("rf_s", "리플럭스 적용 생산 시작", dt.date(2027, 7, 1), "날짜", "", "계획", USER_SRC)
inp("f12_e", "최초 12개월 종료", dt.date(2028, 6, 30), "날짜", "2027.7~2028.6", "계획", USER_SRC)
inp("h_con", "공사 중 정제기 가용시간", None, "h/월", "2027.2~5", "미확인", "-", "0 또는 정상값을 임의 입력하지 않음")
inp("b_trial", "시운전 기간 양품 생산 Batch", None, "Batch/월", "2027.6", "미확인", "-", "")
inp("stab_m", "초기 안정화 기간", None, "개월", "2027.7~", "미확인", "-")

sec("F. 후공정 자원 (정제기 720 h와 별도)")
inp("gb_n", "글로브 박스 수 (5 Gal)", None, "대", "하이닉스·CXMT 공용 여부 확인", "미확인", "-")
inp("gb_h", "글로브 박스 충진 가용 근무시간", None, "h/월", "교대·인력 기준", "미확인", "-", "정제기 720 h를 자동 적용하지 않음")
inp("m200_h", "200 L 충진 (수동/ARS) 가용시간", None, "h/월", "", "미확인", "-")
inp("qc_h", "검사(QC) 가용시간", None, "h/월", "IQC·PQC·FQC·OQC 공용 장비·인력", "미확인", "-")
inp("tank_n", "Product Tank 수", None, "기", "", "미확인", "-", "1기이면 다음 Batch 정제 종료 전 충진 완료 필요")
inp("tank_cap", "Product Tank 용량", None, "kg", "", "미확인", "-")
inp("sc_d", "참고 시나리오: 월 근무일", 22, "일/월", "03 시트 참고 비교용", "추정", "가정 (참고)", "확정값 아님")
inp("sc1", "참고 시나리오 ①: 1교대 일 근무시간", 8, "h/일", "03 시트 참고 비교용", "추정", "가정 (참고)", "확정값 아님")
inp("sc2", "참고 시나리오 ②: 2교대 일 근무시간", 16, "h/일", "03 시트 참고 비교용", "추정", "가정 (참고)", "확정값 아님")
inp("sc3", "참고 시나리오 ③: 연속 근무 일 근무시간", 24, "h/일", "03 시트 참고 비교용 (30일)", "추정", "가정 (참고)", "확정값 아님")

sec("G. 과거 일정 분석 설정")
inp("an_s", "분석 시작 Batch 번호", 39, "#", "6월 첫 Batch", "추정", "분석 설정", "5월(#37·#38)은 일부 기간만 제공 → 기본 제외", "int")
inp("an_e", "분석 종료 Batch 번호", 92, "#", "10월 마지막 Batch", "추정", "분석 설정", "", "int")
inp("roll_k", "롤링 검토 간격 수", 4, "간격", "연속 k간격 평균의 57 h 정합 확인", "추정", "분석 설정", "", "int")
inp("basis", "현재 조건 Batch 간격 기준 선택", 3, "1~4", "1=이론 57 h, 2=과거 최대 11 Batch, 3=과거 일정 관측 추정, 4=26.2톤 역산", "추정", "분석 설정", "04·05 시트 현재 조건 생산 가능량에 적용", "int")
inp("hist_y", "과거 Batch 일정 연도", 2026, "년", "원자료 연도 미표기", "추정", "첨부 PPT 3장이 2026년 출하와 비교", "연도 확인 필요", "int")

sec("H. 재고 · 품질")
inp("inv0", "기초재고 (2026-01-01)", None, "kg", "전 고객 공용 양품 재고", "미확인", "-", "미입력 시 04 시트 재고 계산 보류")
inp("tgt_hx", "하이닉스 목표재고", None, "kg", "하이닉스 우선 확보", "미확인", "-", "월 기말 기준 (매년 반복 가산 금지)")
inp("tgt_all", "전체 목표 기말재고", None, "kg", "", "미확인", "-")
inp("y_all", "정제 Batch 양품률", None, "%", "출하 가능 양품 생산량 = 정제 생산량 × 양품률", "미확인", "-", "미입력 시 100% 상한값으로 표시")
inp("y_hx", "하이닉스 품질 합격률", None, "%", "", "미확인", "-")
inp("y_cx", "CXMT 품질 합격률", None, "%", "", "미확인", "-")
inp("y_ez", "이지켐 품질·색도 합격률", None, "%", "", "미확인", "-")

sec("I. 2027년 출하계획 기준값")
inp("hx27", "하이닉스 2027 연간 출하", 22720, "kg/년", "월별 균등 배분 (=연간÷12)", "계획", USER_SRC, "표시만 반올림", "kg0")
inp("cx27", "CXMT 2027 월 출하", 580, "kg/월", "시작 월부터", "계획", USER_SRC, "29병/월", "kg0")
inp("cx27_s", "CXMT 2027 출하 시작 월", 4, "월", "", "계획", USER_SRC, "4~12월 → 5,220 kg", "int")
inp("ez27", "이지켐 2027 월 출하", 840, "kg/월", "1~12월 전월", "계획", USER_SRC, "6용기/월 → 10,080 kg", "kg0")
last_input_row = r - 1
wsI.auto_filter.ref = f"A5:G{last_input_row}"

# ---- shipment grid
r += 1
grid_title_row = r
wsI.cell(row=r, column=1, value="고객별 월별 출하계획 (kg)").font = font("EB002C", True, size=11)
r += 1
gh = r
header(wsI, gh, ["연도·고객", "연도"] + [f"{m}월" for m in range(1, 13)] + ["연간", "실적/계획 구분", "출처", "비고"])
r += 1
hx26 = [1600, 1420, 1440, 1580, 1200, 1200, 1360, 1420, 1380, 1480, 1480, 1420]
ez26 = [0] * 8 + [560] * 4
SHIP = {}  # (year, cust) -> row

def grid_row(label, year, vals, kind, cls, src, note, key):
    global r
    wsI.cell(row=r, column=1, value=label); style_cell(wsI.cell(row=r, column=1), "text")
    wsI.cell(row=r, column=2, value=year); style_cell(wsI.cell(row=r, column=2), "text", "int")
    for m in range(12):
        c = wsI.cell(row=r, column=3 + m, value=vals[m] if vals else None)
        style_cell(c, kind, "kg")
    c = wsI.cell(row=r, column=15, value=f"=IF(COUNT(C{r}:N{r})=0,\"미제시\",SUM(C{r}:N{r}))"); style_cell(c, "key", "kg")
    for col, v in ((16, cls), (17, src), (18, note)):
        style_cell(wsI.cell(row=r, column=col, value=v), "text")
    SHIP[key] = r
    r += 1

grid_row("하이닉스", 2026, hx26, "in", "원자료 실적/계획 구분 미확인", USER_SRC, "", (2026, "hx"))
grid_row("CXMT", 2026, None, "unk", "미제시", USER_SRC, "2026 CXMT 물량 미제시 — 합계 미포함", (2026, "cx"))
grid_row("이지켐", 2026, ez26, "in", "원자료 실적/계획 구분 미확인", USER_SRC, "9~12월 560 kg (4용기/월)", (2026, "ez"))
# 2027 formulas
rr = r
wsI.cell(row=r, column=1, value="하이닉스"); wsI.cell(row=r, column=2, value=2027)
for m in range(12):
    wsI.cell(row=r, column=3 + m, value=f"={REF['hx27']}/12")
SHIP[(2027, "hx")] = r; r += 1
wsI.cell(row=r, column=1, value="CXMT"); wsI.cell(row=r, column=2, value=2027)
for m in range(12):
    wsI.cell(row=r, column=3 + m, value=f"=IF({m+1}>={REF['cx27_s']},{REF['cx27']},0)")
SHIP[(2027, "cx")] = r; r += 1
wsI.cell(row=r, column=1, value="이지켐"); wsI.cell(row=r, column=2, value=2027)
for m in range(12):
    wsI.cell(row=r, column=3 + m, value=f"={REF['ez27']}")
SHIP[(2027, "ez")] = r; r += 1
for row_, cls, note in ((SHIP[(2027, "hx")], "계획", "22,720 ÷ 12 (표시만 반올림, 연간 합계 유지)"),
                        (SHIP[(2027, "cx")], "계획", "4월부터 580 kg/월"),
                        (SHIP[(2027, "ez")], "계획", "1~12월 전월 840 kg")):
    style_cell(wsI.cell(row=row_, column=1), "text"); style_cell(wsI.cell(row=row_, column=2), "text", "int")
    for m in range(12):
        style_cell(wsI.cell(row=row_, column=3 + m), "calc", "kg")
    c = wsI.cell(row=row_, column=15, value=f"=SUM(C{row_}:N{row_})"); style_cell(c, "key", "kg")
    for col, v in ((16, cls), (17, USER_SRC), (18, note)):
        style_cell(wsI.cell(row=row_, column=col, value=v), "text")
# totals
for yr, lab in ((2026, "2026 제시 합계 (하이닉스+이지켐, CXMT 미포함)"), (2027, "2027 합계 (3개 고객)")):
    rows_ = [SHIP[(yr, k)] for k in ("hx", "cx", "ez")]
    wsI.cell(row=r, column=1, value=lab); style_cell(wsI.cell(row=r, column=1), "text", bold=True)
    wsI.cell(row=r, column=2, value=yr); style_cell(wsI.cell(row=r, column=2), "text", "int")
    for m in range(13):
        col = CL(3 + m)
        c = wsI.cell(row=r, column=3 + m, value="=" + "+".join(f"N({col}{x})" for x in rows_) if m < 12 else f"=SUM(C{r}:N{r})")
        style_cell(c, "key", "kg")
    style_cell(wsI.cell(row=r, column=16, value="합계"), "text")
    style_cell(wsI.cell(row=r, column=17, value="수식"), "text")
    style_cell(wsI.cell(row=r, column=18, value="CXMT 2026 미제시분은 0으로 합산하지 않음(미포함)" if yr == 2026 else "검증: 38,020 kg = 38.02 t"), "text")
    SHIP[(yr, "tot")] = r
    r += 1
# check row
wsI.cell(row=r, column=1, value="검증: 2027 합계 = 38,020 kg"); style_cell(wsI.cell(row=r, column=1), "text")
c = wsI.cell(row=r, column=15, value=f"=IF(ABS(O{SHIP[(2027,'tot')]}-38020)<0.001,\"일치\",\"불일치\")"); style_cell(c, "key")
SHIP["chk"] = r
r += 1
widths(wsI, {"A": 34, "B": 12, "C": 11, "D": 30, "E": 9, "F": 20, "G": 34, "H": 10, "I": 10, "J": 10, "K": 10,
             "L": 10, "M": 10, "N": 10, "O": 12, "P": 18, "Q": 22, "R": 30})
wsI.freeze_panes = "B6"

def ship_ref(year, cust, month):
    return f"'01_Inputs'!${CL(2 + month)}${SHIP[(year, cust)]}"

# ======================================================== 02_Batch_Raw
wsB = wb.create_sheet("02_Batch_Raw")
title(wsB, "02_Batch_Raw | 과거 Batch 생산계획 원자료 전사 · 대조",
      "A~N열 = 원자료(수정 금지) · O열 이후 = 분석/추정 열(수식). 표시일자는 날짜만 있는 생산계획으로, 실제 운전 시각·실적이 아님.")
legend(wsB, 3)
raw_hdr = ["No", "원본 파일", "출처 위치", "Batch 번호", "표시일자\n(발표자 노트)", "실제 시각\n(미확인)", "표시일자 의미\n(확인란)",
           "계획/실적 구분", "원자료 색상 표시\n(PPT 분포도)", "색상 의미\n(확인란)", "발표자 노트\n표기", "분포도 판독\n일(日)",
           "사용자 지시\n목록 일자", "대조 결과"]
ana_hdr = ["월", "간격 분석\n포함(1/0)", "번호 연속", "중복 확인", "이전 Batch와\n간격 (일)", "월 경계", "명목 간격\n(h, 일×24)",
           "최소 가능\n경과 (h)", "최대 가능\n경과 (h)", "롤링 k간격\n누적 (일)", "롤링 k간격 평균\n최대 경과 (h)", "57 h 정합 확인"]
header(wsB, 5, raw_hdr, 1, "hdr", 42)
header(wsB, 5, ana_hdr, 15, "hdr2", 42)

notes_list = [(37, 5, 23), (38, 5, 26), (39, 6, 2), (40, 6, 5), (41, 6, 8), (42, 6, 11), (43, 6, 14), (44, 6, 17), (45, 6, 20),
              (46, 6, 23), (47, 6, 25), (48, 6, 27), (49, 6, 28), (50, 7, 1), (51, 7, 3), (52, 7, 6), (53, 7, 8), (54, 7, 11),
              (55, 7, 14), (56, 7, 17), (57, 7, 20), (58, 7, 23), (59, 7, 26), (60, 7, 29), (61, 8, 1), (62, 8, 4), (63, 8, 7),
              (64, 8, 10), (65, 8, 13), (66, 8, 16), (67, 8, 19), (68, 8, 22), (69, 8, 25), (70, 8, 28), (71, 8, 31), (72, 9, 3),
              (73, 9, 5), (74, 9, 7), (75, 9, 9), (76, 9, 11), (77, 9, 14), (78, 9, 19), (79, 9, 23), (80, 9, 26), (81, 9, 28),
              (82, 10, 1), (83, 10, 4), (84, 10, 7), (85, 10, 10), (86, 10, 12), (87, 10, 15), (88, 10, 18), (89, 10, 20),
              (90, 10, 23), (91, 10, 26), (92, 10, 29)]
# 사용자 지시 목록 (지시문 표) — 별도 전사
user_list = dict(((n, (m, d)) for n, m, d in notes_list))

# 발표자 노트 원문 검증 + 분포도 판독
prs = Presentation(SRC_PPTX)
s3 = prs.slides[2]
notes_txt = s3.notes_slide.notes_text_frame.text
chart = {}
for sh in s3.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().isdigit():
        n = int(sh.text_frame.text.strip()); y = (sh.top + sh.height / 2) / 914400
        if 37 <= n <= 92 and 1.5 < y < 2.75 and abs(sh.width / 914400 - 0.154) < 0.01:
            x = (sh.left + sh.width / 2) / 914400
            color = str(sh.fill.fore_color.rgb) if sh.fill.type == 1 else "?"
            chart[n] = (round(1 + (x - 1.012) / 0.17110), {"FF7900": "주황", "7F7F7F": "회색"}.get(color, color))

B0 = 6
for i, (n, m, d) in enumerate(notes_list):
    rr = B0 + i
    token = f"{m}/{d} #{n}"
    in_notes = token in notes_txt
    vals = [i + 1, SRC_NAME, "3장 발표자 노트 ('원자료 5개 이미지 연결, 중복 제거') · 3장 분포도", n,
            dt.date(2026, m, d), None, None, "생산계획 (원자료) — 실적 미검증", chart.get(n, (None, "?"))[1], None,
            f"{m}/{d}" if in_notes else "노트 미발견", chart.get(n, (None,))[0],
            dt.date(2026, user_list[n][0], user_list[n][1])]
    for j, v in enumerate(vals):
        c = wsB.cell(row=rr, column=1 + j, value=v)
        kind = "in" if j in (3, 4, 8, 10, 11, 12) else ("unk" if j in (5, 6, 9) else "text")
        style_cell(c, kind)
    wsB.cell(row=rr, column=5).number_format = NF["d"]
    wsB.cell(row=rr, column=13).number_format = NF["d"]
    wsB.cell(row=rr, column=14, value=f'=IF(AND(E{rr}=M{rr},DAY(E{rr})=L{rr},K{rr}=MONTH(E{rr})&"/"&DAY(E{rr})),"일치","차이 확인")')
    style_cell(wsB.cell(row=rr, column=14), "calc")
    # analysis
    f = {
        15: f"=MONTH(E{rr})",
        16: f"=IF(AND(D{rr}>{REF['an_s']},D{rr}<={REF['an_e']}),1,0)",
        17: "=\"-\"" if i == 0 else f'=IF(D{rr}-D{rr-1}=1,"연속","확인")',
        18: f'=IF(COUNTIF($D${B0}:$D${B0+len(notes_list)-1},D{rr})>1,"중복","-")',
        19: "" if i == 0 else f"=E{rr}-E{rr-1}",
        20: "" if i == 0 else f'=IF(MONTH(E{rr})<>MONTH(E{rr-1}),"월 경계","")',
        21: "" if i == 0 else f"=S{rr}*24",
        22: "" if i == 0 else f"=MAX(0,(S{rr}-1)*24)",
        23: "" if i == 0 else f"=(S{rr}+1)*24",
        24: f'=IF(ROW()-{B0}>={REF["roll_k"]},E{rr}-INDEX($E${B0}:$E${B0+len(notes_list)-1},ROW()-{B0}+1-{REF["roll_k"]}),"")',
        25: f'=IF(ISNUMBER(X{rr}),(X{rr}+1)*24/{REF["roll_k"]},"")',
        26: f'=IF(ISNUMBER(Y{rr}),IF(Y{rr}<{REF["t_ref"]},"확인 필요: 57 h 미만",""),"")',
    }
    for col, fx in f.items():
        c = wsB.cell(row=rr, column=col, value=fx if fx else None)
        style_cell(c, "calc", {19: "int", 21: "int", 22: "int", 23: "int", 24: "int", 25: "h"}.get(col))
BL = B0 + len(notes_list) - 1
wsB.auto_filter.ref = f"A5:Z{BL}"
wsB.freeze_panes = "E6"
widths(wsB, {"A": 5, "B": 24, "C": 26, "D": 8, "E": 12, "F": 10, "G": 12, "H": 18, "I": 11, "J": 10, "K": 10, "L": 10,
             "M": 12, "N": 10, "O": 6, "P": 9, "Q": 8, "R": 8, "S": 10, "T": 9, "U": 10, "V": 10, "W": 10, "X": 10,
             "Y": 13, "Z": 20})

# monthly summary
S0 = BL + 3
wsB.cell(row=S0 - 1, column=1, value="월별 집계 (번호 개수 vs 원자료 Total) · 간격 분포 — 간격은 뒤 Batch의 월에 귀속, 분석 범위 내").font = font("EB002C", True, 11)
sum_hdr = ["월", "번호 범위", "번호 개수", "원자료 Total", "차이\n(번호−Total)", "번호×190\n(kg, 참고)", "Total×190\n(kg, 참고)",
           "첫 표시일", "마지막 표시일", "간격 수", "평균 간격\n(일)", "중앙값\n(일)", "최소\n(일)", "최대\n(일)",
           "1일", "2일", "3일", "4일", "5일", "6일 이상", "월 경계\n간격 수", "비고"]
header(wsB, S0, sum_hdr, 1, "hdr", 40)
months = [(5, "5월 (일부)", None, "#37~#38만 제공 — 5월 전체 Batch 수로 사용하지 않음"), (6, "6월", 8, ""), (7, "7월", 11, ""),
          (8, "8월", 8, ""), (9, "9월", 8, ""), (10, "10월", 11, "")]
SUMROW = {}
E_ = f"$E${B0}:$E${BL}"; O_ = f"$O${B0}:$O${BL}"; P_ = f"$P${B0}:$P${BL}"; S_ = f"$S${B0}:$S${BL}"; D_ = f"$D${B0}:$D${BL}"; T_ = f"$T${B0}:$T${BL}"
for k, (m, lab, tot, note) in enumerate(months):
    rr = S0 + 1 + k
    SUMROW[m] = rr
    cells = {
        1: lab,
        2: f'="#"&MINIFS_PLACEHOLDER',
        3: f"=COUNTIF({O_},{m})",
        4: tot,
        5: f'=IF(ISNUMBER(D{rr}),C{rr}-D{rr},"원자료 Total 없음")',
        6: f"=C{rr}*{REF['kg_b']}",
        7: f'=IF(ISNUMBER(D{rr}),D{rr}*{REF["kg_b"]},"-")',
        8: f"=_xlfn.MINIFS({E_},{O_},{m})",
        9: f"=_xlfn.MAXIFS({E_},{O_},{m})",
        10: f"=COUNTIFS({O_},{m},{P_},1)",
        11: f'=IF(J{rr}>0,AVERAGEIFS({S_},{O_},{m},{P_},1),"-")',
        12: None,
        13: f'=IF(J{rr}>0,_xlfn.MINIFS({S_},{O_},{m},{P_},1),"-")',
        14: f'=IF(J{rr}>0,_xlfn.MAXIFS({S_},{O_},{m},{P_},1),"-")',
        15: f"=COUNTIFS({O_},{m},{P_},1,{S_},1)",
        16: f"=COUNTIFS({O_},{m},{P_},1,{S_},2)",
        17: f"=COUNTIFS({O_},{m},{P_},1,{S_},3)",
        18: f"=COUNTIFS({O_},{m},{P_},1,{S_},4)",
        19: f"=COUNTIFS({O_},{m},{P_},1,{S_},5)",
        20: f'=COUNTIFS({O_},{m},{P_},1,{S_},">=6")',
        21: f'=COUNTIFS({O_},{m},{P_},1,{T_},"월 경계")',
        22: note,
    }
    cells[2] = f'="#"&_xlfn.MINIFS({D_},{O_},{m})&"~#"&_xlfn.MAXIFS({D_},{O_},{m})'
    for col, v in cells.items():
        if v is None:
            continue
        c = wsB.cell(row=rr, column=col, value=v)
        kind = "in" if col == 4 and tot is not None else ("unk" if col == 4 else ("text" if col in (1, 22) else "calc"))
        style_cell(c, kind, {6: "kg0", 7: "kg0", 8: "md", 9: "md", 11: "b", 12: "b"}.get(col))
    c = wsB.cell(row=rr, column=12)
    wsB[f"L{rr}"] = ArrayFormula(f"L{rr}", f'=IF(J{rr}>0,MEDIAN(IF(({O_}={m})*({P_}=1),{S_})),"-")')
    style_cell(c, "calc", "b")
# total row 6~10
rt = S0 + 1 + len(months)
SUMROW["tot"] = rt
r6, r10 = SUMROW[6], SUMROW[10]
tot_cells = {1: "6~10월 합계", 2: '="#"&' + f'_xlfn.MINIFS({D_},{O_},">=6")&"~#"&_xlfn.MAXIFS({D_},{O_},"<=10")',
             3: f"=SUM(C{r6}:C{r10})", 4: f"=SUM(D{r6}:D{r10})", 5: f"=C{rt}-D{rt}", 6: f"=SUM(F{r6}:F{r10})",
             7: f"=SUM(G{r6}:G{r10})", 8: f"=H{r6}", 9: f"=I{r10}", 10: f"=SUM(J{r6}:J{r10})",
             11: f"=AVERAGEIFS({S_},{P_},1)", 13: f"=_xlfn.MINIFS({S_},{P_},1)", 14: f"=_xlfn.MAXIFS({S_},{P_},1)",
             21: f"=SUM(U{r6}:U{r10})", 22: "5월→6월 경계 간격(7일)은 분석 범위 밖"}
for j, col in enumerate(range(15, 21)):
    tot_cells[col] = f"=SUM({CL(col)}{r6}:{CL(col)}{r10})"
for col, v in tot_cells.items():
    c = wsB.cell(row=rt, column=col, value=v)
    style_cell(c, "text" if col in (1, 22) else "key", {6: "kg0", 7: "kg0", 8: "md", 9: "md", 11: "b"}.get(col), bold=True)
wsB[f"L{rt}"] = ArrayFormula(f"L{rt}", f"=MEDIAN(IF({P_}=1,{S_}))")
style_cell(wsB[f"L{rt}"], "key", "b")
# conditional formats — 집계 차이 · 대조 차이 · 57 h 정합 확인
redfill = PatternFill("solid", fgColor="FFC7CE"); redfont = Font(color="9C0006", bold=True)
wsB.conditional_formatting.add(f"E{r6}:E{rt}", FormulaRule(formula=[f"AND(ISNUMBER(E{r6}),E{r6}<>0)"], fill=redfill, font=redfont))
wsB.conditional_formatting.add(f"C{r6}:D{r10}", FormulaRule(formula=[f"AND(ISNUMBER($D{r6}),$C{r6}<>$D{r6})"], fill=redfill, font=redfont))
wsB.conditional_formatting.add(f"N{B0}:N{BL}", FormulaRule(formula=[f'N{B0}="차이 확인"'], fill=redfill, font=redfont))
wsB.conditional_formatting.add(f"Z{B0}:Z{BL}", FormulaRule(formula=[f'LEFT(Z{B0},2)="확인"'], fill=redfill, font=redfont))
wsB.conditional_formatting.add(f"S{B0}:S{BL}", FormulaRule(formula=[f"AND(ISNUMBER(S{B0}),S{B0}<=2)"], fill=PatternFill("solid", fgColor="FFEB9C")))
# check rows
rc = rt + 2
checks = [("대조 결과 '차이 확인' 건수 (노트·분포도·사용자 목록)", f'=COUNTIF($N${B0}:$N${BL},"차이 확인")'),
          ("번호 불연속 건수", f'=COUNTIF($Q${B0}:$Q${BL},"확인")'),
          ("중복 Batch 번호 건수", f'=COUNTIF($R${B0}:$R${BL},"중복")'),
          ("월별 Total ≠ 번호 개수 월 수 (6~10월)", f'=SUMPRODUCT(--(E{r6}:E{r10}<>0))'),
          ("57 h 정합 확인 필요 롤링 구간 수", f'=COUNTIF($Z${B0}:$Z${BL},"확인 필요*")')]
CHK = {}
for k, (lab, fx) in enumerate(checks):
    wsB.cell(row=rc + k, column=1, value=lab); style_cell(wsB.cell(row=rc + k, column=1), "text")
    wsB.merge_cells(start_row=rc + k, start_column=1, end_row=rc + k, end_column=4)
    c = wsB.cell(row=rc + k, column=5, value=fx); style_cell(c, "key", "int")
    CHK[k] = f"'02_Batch_Raw'!$E${rc+k}"
rn = rc + len(checks) + 1
for k, t in enumerate([
    "※ 원본 이미지 5개는 첨부 PPT에 포함되어 있지 않음 → 발표자 노트 전사값과 3장 분포도(x좌표 → 일자 판독), 사용자 지시 목록 3중 대조. 원본 이미지 대조는 추가 확인사항.",
    "※ 원자료 색상(주황/회색)은 PPT 분포도 표시색이며 원본 이미지 색상 의미(고객·품질·상태)는 미확인 — 임의 해석하지 않음.",
    "※ 날짜만 있는 자료: D일 간격의 실제 경과시간은 (D−1)×24 ~ (D+1)×24 h 범위. 2일 간격을 48 h로 단정하지 않음.",
    "※ 번호 개수 ≠ 양품 완료 Batch 수. 원자료 Total의 집계 대상(완료/양품/출하)과 작성 시점 확인 필요."]):
    c = wsB.cell(row=rn + k, column=1, value=t); c.font = font(GRAY, italic=True)

# ======================================================== 03_Capacity_Model
wsC = wb.create_sheet("03_Capacity_Model")
title(wsC, "03_Capacity_Model | 정제기 · 후공정 병렬 운영 생산능력 계산",
      "정제기(공용 1대)는 앞 Batch 검사·충진 중에도 다음 Batch를 정제. 한 Batch Lead Time을 정제 생산 간격으로 사용하지 않음.")
legend(wsC, 3)
widths(wsC, {"A": 44, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 44})
C = {}   # key -> absolute ref
row = 5

def cblock(label):
    global row
    row += 1
    c = wsC.cell(row=row, column=1, value=label); c.font = font("EB002C", True, size=11)
    row += 1

def cline(key, label, fx, unit, basis, kind="calc", nf="h2"):
    global row
    style_cell(wsC.cell(row=row, column=1, value=label), "text")
    c = wsC.cell(row=row, column=2, value=fx); style_cell(c, kind, nf)
    style_cell(wsC.cell(row=row, column=3, value=unit), "text")
    wsC.merge_cells(start_row=row, start_column=4, end_row=row, end_column=8)
    style_cell(wsC.cell(row=row, column=4, value=basis), "note")
    if key:
        C[key] = f"'03_Capacity_Model'!$B${row}"
    row += 1

I = REF
# --- Block 1: Lead Time
cblock("1. 충진 경로별 한 Batch 제조 Lead Time (확인된 공정시간 합계) — 정제 생산 간격 아님")
header(wsC, row, ["공정 단계", "5 Gal 경로\n(하이닉스·CXMT, h)", "200 L 경로\n(이지켐, h)", "적용 단위", "", "", "", "비고"])
row += 1
lt_start = row
steps = [("입고 · IQC", "t_iqc", "t_iqc", "원료 Lot당 (공통 적용 시 Batch별 중복 배정 금지)"),
         ("준비 · 투입", "t_prep", "t_prep", "57 h 포함 여부 확인"),
         ("정제 (공용 정제기 1대)", "t_ref", "t_ref", "Batch당 — 포함 범위 확인"),
         ("PQC", "t_pqc", "t_pqc", "Batch당"),
         ("Product Tank 이송", "t_tank", "t_tank", "Batch당"),
         ("FQC", "t_fqc", "t_fqc", "Batch당 (대기 별도)"),
         ("FQC 대기", "t_fqcw", "t_fqcw", "Batch당"),
         ("고객별 충진 (필터 포함)", "t_5g_b", "t_200", "5 Gal: 약 8병 16 h · 200 L: 8 h 1회 적용 (용기/Batch 단위 확인)"),
         ("OQC · 출하 작업", "t_oqc", "t_oqc", "출하 Lot당 · 운송시간 제외")]
for lab, k1, k2, unit in steps:
    style_cell(wsC.cell(row=row, column=1, value=lab), "text")
    for col, k in ((2, k1), (3, k2)):
        c = wsC.cell(row=row, column=col, value=f'=IF(ISNUMBER({I[k]}),{I[k]},"미확인")'); style_cell(c, "link", "h")
    wsC.merge_cells(start_row=row, start_column=4, end_row=row, end_column=7)
    style_cell(wsC.cell(row=row, column=4, value=unit), "note")
    row += 1
lt_end = row - 1
style_cell(wsC.cell(row=row, column=1, value="확인된 공정시간 합계 (미확인 항목 제외 → 하한값)"), "text", bold=True)
for col in (2, 3):
    c = wsC.cell(row=row, column=col, value=f"=SUM({CL(col)}{lt_start}:{CL(col)}{lt_end})"); style_cell(c, "key", "h")
C["lt_5g"] = f"'03_Capacity_Model'!$B${row}"; C["lt_200"] = f"'03_Capacity_Model'!$C${row}"
wsC.merge_cells(start_row=row, start_column=4, end_row=row, end_column=8)
style_cell(wsC.cell(row=row, column=4, value="한 Batch가 입고→출하까지 걸리는 확인분 시간. 다음 Batch 정제 시작 간격으로 사용하지 않음."), "note")
row += 1
style_cell(wsC.cell(row=row, column=1, value="미확인 항목 수"), "text")
for col in (2, 3):
    c = wsC.cell(row=row, column=col, value=f'=COUNTIF({CL(col)}{lt_start}:{CL(col)}{lt_end},"미확인")'); style_cell(c, "calc", "int")
C["lt_unk"] = f"'03_Capacity_Model'!$B${row}"
row += 1

# --- Block 2: 이론
cblock("2. ① 정제시간만 적용한 이론 생산능력 (정제기 연속 가동 상한)")
cline("th_mh", "월 비교 기준 시간", f"={I['mh']}", "h/월", "30일×24 h", "link", "h")
cline("th_t", "Batch당 정제시간", f"={I['t_ref']}", "h/Batch", "현재 103℃ 제공값", "link", "h")
cline("th_bm", "이론 Batch/월 (소수)", f"={C['th_mh']}/{C['th_t']}", "Batch/월", "720 ÷ 57")
cline("th_bi", "월내 완료 가능 정수 Batch", f"=INT({C['th_bm']})", "Batch/월", "INT(720 ÷ 57)", "key", "int")
cline("th_carry_h", "월말 진행 중 Batch 경과시간 (다음 월 이월)", f"={C['th_mh']}-{C['th_bi']}*{C['th_t']}", "h", "720 − 정수 Batch × 57", nf="h")
cline("th_carry_b", "월말 진행 중 Batch 진행률", f"={C['th_carry_h']}/{C['th_t']}", "Batch", "이월분 — 다음 월 완료")
cline("th_kgm", "월 생산량 (정수 Batch × 190 kg)", f"={C['th_bi']}*{I['kg_b']}", "kg/월", "", "key", "kg0")
cline("th_by", "연간 Batch (달력 연속 운전, 월간 이월 반영)", f"=INT({I['yh']}/{C['th_t']})", "Batch/년", "INT(8,760 ÷ 57) — 정지·보수 미반영 상한", "calc", "int")
cline("th_kgy", "연간 생산량 (달력 연속)", f"={C['th_by']}*{I['kg_b']}", "kg/년", "", "key", "kg0")
cline("th_by12", "참고: 단순 연간 환산 (월 정수 Batch × 12)", f"={C['th_bi']}*12", "Batch/년", "월간 이월 미반영 — 참고값", "calc", "int")
cline("th_kgy12", "참고: 단순 연간 환산 생산량", f"={C['th_by12']}*{I['kg_b']}", "kg/년", "", "calc", "kg0")

# --- Block 3: 과거 최대 11
cblock("3. ② 과거 월 최대 11 Batch 기준")
cline("mx_b", "과거 월 최대 Batch", f"={I['max11']}", "Batch/월", "히스토리 최대 (제공)", "link", "int")
cline("mx_ref", "11 Batch 정제시간 합계", f"={C['mx_b']}*{C['th_t']}", "h/월", "11 × 57", nf="h")
cline("mx_gap", "월 720 h와의 차이", f"={C['th_mh']}-{C['mx_ref']}", "h/월", "여유·전환·대기·보수·월 경계 가능성 — 손실/충진시간으로 단정하지 않음", nf="h")
cline("mx_int", "월 시간을 11 Batch에 배분한 실효시간 (30일 월)", f"={C['th_mh']}/{C['mx_b']}", "h/Batch", "720 ÷ 11 — 참고값", "key", "h2")
cline("mx_int31", "참고: 31일 월 기준 실효시간", f"=31*24/{C['mx_b']}", "h/Batch", "744 ÷ 11 — 과거 11 Batch 월(7·8·10월)은 모두 31일 월", nf="h2")
cline("mx_diff", "실효시간 − 정제시간 (30일 월)", f"={C['mx_int']}-{C['th_t']}", "h/Batch", "비정제 시간 상당 (구성 미확인)", nf="h2")
cline("mx_kgm", "월 생산량", f"={C['mx_b']}*{I['kg_b']}", "kg/월", "11 × 190", "key", "kg0")
cline("mx_by12", "단순 연간 환산 (11 × 12)", f"={C['mx_b']}*12", "Batch/년", "참고값", "calc", "int")
cline("mx_kgy12", "단순 연간 환산 생산량", f"={C['mx_by12']}*{I['kg_b']}", "kg/년", "", "calc", "kg0")
cline("mx_by", "달력 연속 환산 (8,760 ÷ 실효시간)", f"=INT({I['yh']}/{C['mx_int']})", "Batch/년", "월간 이월 반영", "calc", "int")
cline("mx_kgy", "달력 연속 환산 생산량", f"={C['mx_by']}*{I['kg_b']}", "kg/년", "", "key", "kg0")

# --- Block 4: 관측 추정
cblock("4. ③ 과거 일정 기반 실효 Batch 간격 추정 (02_Batch_Raw)")
Bref = "'02_Batch_Raw'!"
E_r = f"{Bref}$E${B0}:$E${BL}"; D_r = f"{Bref}$D${B0}:$D${BL}"; P_r = f"{Bref}$P${B0}:$P${BL}"; S_r = f"{Bref}$S${B0}:$S${BL}"
cline("ob_range", "사용 데이터 범위", f'="#"&{I["an_s"]}&" ~ #"&{I["an_e"]}&" (표시일자, 생산계획)"', "", "기본: 6~10월 (#39~#92). 5월은 일부만 제공되어 제외", "link", None)
cline("ob_d0", "분석 시작 표시일자", f'=INDEX({E_r},MATCH({I["an_s"]},{D_r},0))', "날짜", "", "link", "yyyy-mm-dd")
cline("ob_d1", "분석 종료 표시일자", f'=INDEX({E_r},MATCH({I["an_e"]},{D_r},0))', "날짜", "", "link", "yyyy-mm-dd")
cline("ob_days", "달력 경과일수", f"={C['ob_d1']}-{C['ob_d0']}", "일", "", nf="int")
cline("ob_n", "Batch 간격 수", f"={I['an_e']}-{I['an_s']}", "간격", "번호 연속 가정 (02 시트 번호 연속 확인)", nf="int")
cline("ob_ncheck", "검증: 분석 포함 간격 수 (02 시트)", f"=SUM({P_r})", "간격", "위 값과 같아야 함", nf="int")
cline("ob_avg_d", "평균 간격", f"={C['ob_days']}/{C['ob_n']}", "일/Batch", "개별 간격 평균과 동일 (누적 방식)", nf="0.000")
cline("ob_int", "추정 실효 Batch 간격 (중앙 추정)", f"={C['ob_days']}*24/{C['ob_n']}", "h/Batch", "경과일수×24 ÷ 간격 수 — 표시일자가 매 Batch 동일 이벤트라고 가정", "key", "h2")
cline("ob_lo", "추정 범위 하한 (날짜 해상도)", f"=({C['ob_days']}-1)*24/{C['ob_n']}", "h/Batch", "시각 미상 → 양 끝 ±1일 → 장기 누적이라 범위가 좁음", nf="h2")
cline("ob_hi", "추정 범위 상한 (날짜 해상도)", f"=({C['ob_days']}+1)*24/{C['ob_n']}", "h/Batch", "", nf="h2")
cline("ob_med", "개별 간격 중앙값", f"='02_Batch_Raw'!$L${SUMROW['tot']}*24", "h", "3일 = 72 h (명목)", "link", "h")
cline("ob_share3", "3일 간격 비중", f"='02_Batch_Raw'!$Q${SUMROW['tot']}/'02_Batch_Raw'!$J${SUMROW['tot']}", "%", "", "link", "pct")
cline("ob_ovh", "실효 간격 − 정제 57 h (비정제 시간 추정)", f"={C['ob_int']}-{C['th_t']}", "h/Batch", "준비·투입, 세척·전환, 검사·Tank 대기, 계획 여유 등 가능성 — 구성 미확인", "key", "h2")
cline("ob_ovh_lo", "비정제 시간 범위 하한", f"={C['ob_lo']}-{C['th_t']}", "h/Batch", "", nf="h2")
cline("ob_ovh_hi", "비정제 시간 범위 상한", f"={C['ob_hi']}-{C['th_t']}", "h/Batch", "", nf="h2")
cline("ob_b30", "월 Batch (30일, 소수)", f"=30*24/{C['ob_int']}", "Batch/월", "", nf="b")
cline("ob_b31", "월 Batch (31일, 소수)", f"=31*24/{C['ob_int']}", "Batch/월", "과거 11 Batch 월(31일)과 비교", nf="b")
cline("ob_by", "연간 Batch (달력 연속, 정수)", f"=INT({I['yh']}/{C['ob_int']})", "Batch/년", "", nf="int")
cline("ob_kgy", "연간 생산량 (달력 연속)", f"={C['ob_by']}*{I['kg_b']}", "kg/년", "현재 103℃ 조건 · 정지·보수 추가 미반영", "key", "kg0")
cline("ob_kgm30", "월 생산량 참고 (30일, 소수 Batch)", f"={C['ob_b30']}*{I['kg_b']}", "kg/월", "", nf="kg0")
cline("ob_alt37", "참고: #37 포함 시 평균 간격", f"=({Bref}$E${BL}-{Bref}$E${B0})*24/({Bref}$D${BL}-{Bref}$D${B0})", "h/Batch", "5/26→6/2 (7일) 포함 — 자료 공백 가능성", nf="h2")
tot_r = SUMROW["tot"]
cline("ob_tot_days", "참고: 6~10월 달력일수", "=DATE(" + I["hist_y"] + ",11,1)-DATE(" + I["hist_y"] + ",6,1)", "일", "", nf="int")
cline("ob_tot_int", "참고: 원자료 Total(46)을 완료 Batch로 해석 시 간격", f"={C['ob_tot_days']}*24/'02_Batch_Raw'!$D${tot_r}", "h/Batch", "Total 집계 대상 미확인 — 참고값", nf="h2")
cline("ob_num_int", "참고: 번호 개수(54)를 월 달력으로 나눈 간격", f"={C['ob_tot_days']}*24/'02_Batch_Raw'!$C${tot_r}", "h/Batch", "월 경계 이월 포함", nf="h2")
cline("ob_dense", "57 h 정합 확인 필요 롤링 구간 수", f"={CHK[4]}", "구간", "연속 k간격 평균 최대 경과 < 57 h (6/23~7/3, 9/3~9/11)", "link", "int")

# --- Block 5: 26.2 t
cblock("5. 현재 Capa. 26.2톤/년 환산 관계 (기준값 유지 · 조정하지 않음)")
cline("cp_kg", "현재 연간 Capa.", f"={I['capa_now']}*1000", "kg/년", "제공 기준값", "link", "kg0")
cline("cp_b", "Batch 상당 (소수)", f"={C['cp_kg']}/{I['kg_b']}", "Batch/년", "26,200 ÷ 190", nf="0.000")
cline("cp_bint", "정수 Batch (내림)", f"=INT({C['cp_b']})", "Batch/년", "", nf="int")
cline("cp_kgint", "정수 Batch (내림) 생산량", f"={C['cp_bint']}*{I['kg_b']}", "kg/년", "137 × 190", nf="kg0")
cline("cp_bround", "정수 Batch (반올림)", f"=ROUND({C['cp_b']},0)", "Batch/년", "", nf="int")
cline("cp_kground", "정수 Batch (반올림) 생산량", f"={C['cp_bround']}*{I['kg_b']}", "kg/년", "138 × 190 = 26,220 kg → 26.22톤 (표시 26.2톤)", nf="kg0")
cline("cp_bm", "월 평균 Batch 상당", f"={C['cp_b']}/12", "Batch/월", "실제 월 계획 아님", nf="b")
cline("cp_int", "26.2톤 역산 평균 Batch 간격 (8,760 h 연속 가정)", f"={I['yh']}/{C['cp_b']}", "h/Batch", "참고 역산값", "key", "h2")
cline("cp_vs_ob", "관측 추정 연 생산량 − 26.2톤", f"=({C['ob_kgy']}-{C['cp_kg']})/1000", "t/년", "차이 원인(가동일·Batch량·집계 기준) 확인", nf="t")

# --- Block 6: 비교표
cblock("6. 생산능력 비교표 (정제기 기준, 190 kg/Batch)")
header(wsC, row, ["구분", "Batch 간격\n(h)", "월 Batch\n(30일)", "월 Batch\n(31일)", "연 Batch\n(달력 연속, 정수)", "연 생산량\n(t)", "26.2톤 대비\n(t)", "근거 · 적용 범위"])
row += 1
CMP0 = row
cmp_rows = [("① 이론 (정제 57 h만)", f"={C['th_t']}", "상한 — 정지·전환·대기 0 가정"),
            ("현재 Capa. 26.2톤 역산", f"={C['cp_int']}", "제공 기준값의 등가 간격 (역산)"),
            ("② 과거 월 최대 11 Batch", f"={C['mx_int']}", "30일 월 기준 배분 실효시간"),
            ("③ 과거 일정 관측 추정", f"={C['ob_int']}", "2026.6~10 계획 일정 #39~#92"),
            ("참고: 원자료 Total 기준", f"={C['ob_tot_int']}", "Total=완료 Batch 해석 시 (미확인)")]
for lab, fx, basis in cmp_rows:
    style_cell(wsC.cell(row=row, column=1, value=lab), "text", bold=True)
    vals = [fx, f"=30*24/B{row}", f"=31*24/B{row}", f"=INT({I['yh']}/B{row})", f"=E{row}*{I['kg_b']}/1000", f"=F{row}-{I['capa_now']}"]
    for j, v in enumerate(vals):
        c = wsC.cell(row=row, column=2 + j, value=v)
        style_cell(c, "key" if j == 4 else "calc", ["h2", "b", "b", "int", "t", "t"][j])
    style_cell(wsC.cell(row=row, column=8, value=basis), "note")
    row += 1
C["cmp0"] = CMP0
# selected basis
cline("basis_int", "04·05 시트 적용 현재 조건 Batch 간격 (01_Inputs 선택)",
      f"=CHOOSE({I['basis']},{C['th_t']},{C['mx_int']},{C['ob_int']},{C['cp_int']})", "h/Batch",
      "1=이론, 2=과거 최대, 3=관측 추정(기본), 4=26.2톤 역산", "key", "h2")
cline("basis_name", "선택 기준명", f'=CHOOSE({I["basis"]},"이론 57 h","과거 최대 11 Batch","과거 일정 관측 추정","26.2톤 역산")', "", "", "calc", None)

# --- Block 7: 후공정
cblock("7. ④ 후공정 작업량 vs 가용시간 (정제기 720 h와 별도 · 참고 시나리오는 확정값 아님)")
cline("pp_b_5g", "Batch당 5 Gal 충진 작업 (약 8병)", f"={I['t_5g_b']}", "h/Batch", "제공 작업 기준", "link", "h")
cline("pp_b_fqc", "Batch당 FQC", f"={I['t_fqc']}", "h/Batch", "", "link", "h")
cline("pp_gb_need26", "2026 월 최대 5 Gal 충진 작업시간", "='04_Monthly_2026_2027'!$Z$32", "h/월", "04 시트 2026 최대", "link", "h")
cline("pp_gb_need27", "2027 월 최대 5 Gal 충진 작업시간", "='04_Monthly_2026_2027'!$Z$33", "h/월", "04 시트 2027 최대 (하이닉스+CXMT)", "link", "h")
cline("pp_200_need27", "2027 월 최대 200 L 충진 작업시간", "='04_Monthly_2026_2027'!$AD$33", "h/월", "ARS 시간 미확인 시 수동 8 h/용기 참고", "link", "h")
cline("pp_qc_need", "검사·출하 작업시간 (Batch당 IQC+FQC+OQC 상한)", f"='04_Monthly_2026_2027'!$AE$32", "h/월", "2026 월 최대 — 상한 가정", "link", "h")
header(wsC, row, ["후공정 가용시간 시나리오", "일 근무\n(h/일)", "월 근무일\n(일)", "월 가용\n(h)", "2026 5 Gal\n최대 부하율", "2027 5 Gal\n최대 부하율", "2027 5 Gal\n여유 (h)", "비고"])
row += 1
PP0 = row
for lab, hk, dk, note in (("입력값 (글로브 박스 가용 근무시간)", None, None, "01_Inputs F — 미확인 시 '미확인'"),
                          ("참고 ① 1교대", "sc1", "sc_d", "가정"), ("참고 ② 2교대", "sc2", "sc_d", "가정"),
                          ("참고 ③ 연속 근무 (30일)", "sc3", None, "가정 — 정제기와 동일 720 h 가정 아님 확인용")):
    style_cell(wsC.cell(row=row, column=1, value=lab), "text", bold=True)
    if hk is None:
        vals = ["-", "-", f'=IF(ISNUMBER({I["gb_h"]}),{I["gb_h"]},"미확인")']
    else:
        vals = [f"={I[hk]}", f"={I[dk]}" if dk else "=30", f"=B{row}*C{row}"]
    vals += [f'=IF(ISNUMBER(D{row}),{C["pp_gb_need26"]}/D{row},"미확인")', f'=IF(ISNUMBER(D{row}),{C["pp_gb_need27"]}/D{row},"미확인")',
             f'=IF(ISNUMBER(D{row}),D{row}-{C["pp_gb_need27"]},"미확인")']
    for j, v in enumerate(vals):
        c = wsC.cell(row=row, column=2 + j, value=v); style_cell(c, "calc", ["h", "int", "h", "pct", "pct", "h"][j])
    style_cell(wsC.cell(row=row, column=8, value=note), "note")
    row += 1
C["pp0"] = PP0

# --- Block 8: Tank 회전 / 병렬
cblock("8. ⑤ 병렬 운영 병목 확인 — Product Tank 1기 가정 시 Tank 회전 (충진 완료 ≤ 다음 Batch 정제 종료)")
cline("tk_occ", "정제 후 Tank 점유 확인분 (PQC+이송+FQC+대기, 충진 제외)",
      f"=N({I['t_pqc']})+N({I['t_tank']})+{I['t_fqc']}+N({I['t_fqcw']})", "h", "미확인 항목은 미반영 → 하한", nf="h")
cline("rf_eq_int", "참고: 47.2톤 달성 등가 Batch 간격 (190 kg 유지 · 8,760 h)", f"={I['yh']}/({I['capa_rf']}*1000/{I['kg_b']})", "h/Batch", "역산 참고 — 리플럭스 후 실제 정제시간 미확인", nf="h2")
header(wsC, row, ["5 Gal Batch 충진 (16 h 작업) 경과시간", "일 근무\n(h/일)", "충진 경과\n(h)", "Tank 점유\n합계 (h)", "여유 vs 57 h\n(이론 간격)", "여유 vs 관측\n실효 간격", "여유 vs 47.2톤\n등가 간격", "판단"])
row += 1
TK0 = row
for lab, hk in (("참고 ① 1교대", "sc1"), ("참고 ② 2교대", "sc2"), ("참고 ③ 연속", "sc3")):
    style_cell(wsC.cell(row=row, column=1, value=lab), "text", bold=True)
    vals = [f"={I[hk]}", f"={I['t_5g_b']}/(B{row}/24)", f"={C['tk_occ']}+C{row}", f"={C['th_t']}-D{row}",
            f"={C['ob_int']}-D{row}", f"={C['rf_eq_int']}-D{row}",
            f'=IF(G{row}<0,"리플럭스 후 병목 가능",IF(E{row}<10,"여유 작음","여유"))']
    for j, v in enumerate(vals):
        c = wsC.cell(row=row, column=2 + j, value=v); style_cell(c, "calc", ["h", "h", "h", "h", "h", "h", None][j])
    row += 1
C["tk0"] = TK0
style_cell(wsC.cell(row=row, column=1, value="※ Tank 수·용량이 확인되면(01_Inputs F) 2기 이상일 때 이 제약은 완화됨. 충진시간을 정제시간에 더해 정제기 능력을 줄이지 않음."), "note")
row += 1

# --- Block 9: Gantt
cblock("9. 병렬 운영 Gantt (시간 h) — 정제기 1대 연속 · 검사/충진은 다음 Batch 정제와 겹침")
cline("g_int", "Gantt 적용 Batch 간격", f"=ROUND({C['ob_int']},1)", "h/Batch", "관측 추정값 (이론 57 h로 보려면 이 셀을 57로 변경)", "est", "h")
header(wsC, row, ["작업", "시작 (h)", "종료 (h)", "자원", "", "", "", "비고"])
row += 1
G0 = row
gantt = []
gi = C["g_int"]
for b in range(1, 4):
    rs = f"=({b}-1)*{gi}"
    gantt.append((f"Batch {b} 정제", rs, f"=B{{r}}+{I['t_ref']}", "공용 정제기", "정제 57 h"))
    gantt.append((f"Batch {b} 비정제 (추정)", f"=B{{prev}}+{I['t_ref']}", f"=({b})*{gi}", "공용 정제기", "준비·세척·대기 등 (구성 미확인)"))
    gantt.append((f"Batch {b} PQC·이송·FQC", f"=C{{pp}}", f"=B{{r}}+{C['tk_occ']}", "검사·Tank", "확인분 2 h (미확인 항목 미반영)"))
    gantt.append((f"Batch {b} 5 Gal 충진 (8병)", f"=C{{prev}}", f"=B{{r}}+{I['t_5g_b']}", "글로브 박스", "연속 작업 가정 16 h, 필터 포함"))
    gantt.append((f"Batch {b} OQC·출하", f"=C{{prev}}", f"=B{{r}}+{I['t_oqc']}", "검사·출하", ""))
for k, (lab, s, e, res, note) in enumerate(gantt):
    rr = G0 + k
    blk = k // 5; pos = k % 5
    ref_rows = {"r": rr, "prev": rr - 1, "pp": G0 + blk * 5}
    s2 = s.replace("{prev}", str(rr - 1)).replace("{pp}", str(G0 + blk * 5)).replace("{r}", str(rr))
    e2 = e.replace("{prev}", str(rr - 1)).replace("{pp}", str(G0 + blk * 5)).replace("{r}", str(rr))
    style_cell(wsC.cell(row=rr, column=1, value=lab), "text")
    style_cell(wsC.cell(row=rr, column=2, value=s2), "calc", "h")
    style_cell(wsC.cell(row=rr, column=3, value=e2), "calc", "h")
    style_cell(wsC.cell(row=rr, column=4, value=res), "text")
    style_cell(wsC.cell(row=rr, column=8, value=note), "note")
GN = G0 + len(gantt) - 1
C["g0"], C["gn"] = G0, GN
# grid 0..216 step 6 in columns J..
GRID0 = 10
for j in range(0, 37):
    c = wsC.cell(row=G0 - 1, column=GRID0 + j, value=j * 6); c.font = font(GRAY, size=7); c.alignment = CENTER
    wsC.column_dimensions[CL(GRID0 + j)].width = 3.2
    for rr in range(G0, GN + 1):
        cc = wsC.cell(row=rr, column=GRID0 + j, value=f'=IF(AND({CL(GRID0+j)}${G0-1}+6>$B{rr},{CL(GRID0+j)}${G0-1}<$C{rr}),1,"")')
        cc.font = font("FFFFFF", size=6); cc.border = Border(left=Side(style="hair", color="D9D9D9"))
rng = f"{CL(GRID0)}{G0}:{CL(GRID0+36)}{GN}"
for kw, col in (("정제", "EB002C"), ("비정제", "BFBFBF"), ("FQC", "2E75B6"), ("충진", "70AD47"), ("OQC", "7F7F7F")):
    wsC.conditional_formatting.add(rng, FormulaRule(formula=[f'AND({CL(GRID0)}{G0}=1,ISNUMBER(SEARCH("{kw}",$A{G0})))'],
                                                   fill=PatternFill("solid", fgColor=col), font=Font(color=col)))
wsC.cell(row=G0 - 1, column=GRID0 - 1, value="h →").font = font(GRAY, size=7)
wsC.freeze_panes = "B5"

# ======================================================== 04_Monthly_2026_2027
wsM = wb.create_sheet("04_Monthly_2026_2027")
title(wsM, "04_Monthly_2026_2027 | 월별 생산 가능량 · 고객별 출하 · 충진 작업 · 재고 (24개월)",
      "생산 가능 Batch = INT((가용시간 + 이월시간) ÷ 적용 Batch 간격). 미확인 조건은 '미확인/보류'로 표시 — 0으로 치환하지 않음.")
legend(wsM, 3)
mh = ["연월", "연도", "월 일수", "운전 조건", "달력시간\n(h)", "정지·보수\n입력 (h)", "정제기\n가용시간 (h)", "적용 Batch\n간격 (h)",
      "이월 입력\n(h)", "생산 가능\n완료 Batch", "월말 진행 중\n이월 (h)", "과거계획\n번호 개수", "과거계획\n원자료 Total",
      "확정 생산 Batch\n(실적 입력)", "적용 Batch", "적용 기준", "Batch당\n생산량 (kg)", "정제 생산량\n(kg)",
      "출하 가능 양품\n생산량 (kg)", "하이닉스\n출하 (kg)", "CXMT\n출하 (kg)", "이지켐\n출하 (kg)", "출하 합계\n(kg)",
      "필요 Batch\n(출하÷Batch량)", "5 Gal 병수\n(병)", "5 Gal 충진\n작업 (h)", "200 L\n용기 수", "200 L 단위\n시간 (h)", "200 L\n충진 방식",
      "200 L 충진\n작업 (h)", "검사·출하 작업\n(h, 상한)", "글로브 박스\n가용 (h)", "글로브 박스\n여유 (h)",
      "당월 양품 생산\n−출하 (kg)", "누적 증감 (kg)\n기초재고 제외", "기초재고\n(kg)", "공급 가능량\n(kg)", "① 하이닉스\n배정 (kg)",
      "② 하이닉스\n목표재고 확보", "③ CXMT\n배정 (kg)", "④ 이지켐\n배정 (kg)", "출하 부족\n(kg)", "기말재고\n(kg)",
      "목표\n기말재고 (kg)", "필요 생산량\n(kg)", "여유·부족\n(기말−목표)", "계산 가능 여부", "미확인 항목"]
header(wsM, 5, mh, 1, "hdr", 54)
M0 = 6
for k in range(24):
    rr = M0 + k
    y = 2026 + k // 12; m = k % 12 + 1
    first = k == 0
    p = rr - 1
    f = {}
    f["A"] = dt.date(y, m, 1)
    f["B"] = f"=YEAR(A{rr})"
    f["C"] = f"=DAY(EOMONTH(A{rr},0))"
    f["D"] = (f'=IF(A{rr}<{I["con_s"]},"현재(103℃)",IF(A{rr}<={I["con_e"]},"리플럭스 공사",'
              f'IF(A{rr}<{I["rf_s"]},"시운전","리플럭스 적용")))')
    f["E"] = f"=C{rr}*24"
    f["F"] = None
    f["G"] = (f'=IF(D{rr}="리플럭스 공사",IF(ISNUMBER({I["h_con"]}),{I["h_con"]},"미확인"),'
              f'IF(D{rr}="시운전","시운전",E{rr}-N(F{rr})))')
    f["H"] = (f'=IF(OR(D{rr}="현재(103℃)",D{rr}="리플럭스 공사"),{C["basis_int"]},IF(D{rr}="리플럭스 적용",'
              f'IF(ISNUMBER({I["t_ref_rf"]}),{I["t_ref_rf"]}+IF(ISNUMBER({I["ovh_rf"]}),{I["ovh_rf"]},{C["ob_ovh"]}),"미확인"),"-"))')
    f["I"] = "=0" if first else f"=IF(AND(ISNUMBER(K{p}),D{p}=D{rr}),K{p},0)"
    f["J"] = (f'=IF(D{rr}="시운전",IF(ISNUMBER({I["b_trial"]}),{I["b_trial"]},"미확인"),'
              f'IF(AND(ISNUMBER(G{rr}),ISNUMBER(H{rr})),INT((G{rr}+I{rr})/H{rr}),"미확인"))')
    f["K"] = f'=IF(AND(ISNUMBER(G{rr}),ISNUMBER(H{rr}),ISNUMBER(J{rr})),G{rr}+I{rr}-J{rr}*H{rr},"")'
    f["L"] = (f'=IF(COUNTIFS({E_r},">="&A{rr},{E_r},"<="&EOMONTH(A{rr},0))=0,"",'
              f'COUNTIFS({E_r},">="&A{rr},{E_r},"<="&EOMONTH(A{rr},0)))')
    if y == 2026 and m in SUMROW:
        f["M"] = f"=IF(ISNUMBER('02_Batch_Raw'!$D${SUMROW[m]}),'02_Batch_Raw'!$D${SUMROW[m]},\"\")"
    else:
        f["M"] = '=""'
    f["N"] = None
    f["O"] = f'=IF(ISNUMBER(N{rr}),N{rr},J{rr})'
    f["P"] = f'=IF(ISNUMBER(N{rr}),"실적 입력",IF(ISNUMBER(J{rr}),"추정 생산 가능량","미확인"))'
    f["Q"] = f'=IF(D{rr}="리플럭스 적용",IF(ISNUMBER({I["kg_b_rf"]}),{I["kg_b_rf"]},{I["kg_b"]}),{I["kg_b"]})'
    f["R"] = f'=IF(ISNUMBER(O{rr}),O{rr}*Q{rr},"미확인")'
    f["S"] = f'=IF(ISNUMBER(R{rr}),R{rr}*IF(ISNUMBER({I["y_all"]}),{I["y_all"]},1),"미확인")'
    f["T"] = f"={ship_ref(y, 'hx', m)}"
    f["U"] = f'=IF(ISNUMBER({ship_ref(y, "cx", m)}),{ship_ref(y, "cx", m)},"미제시")'
    f["V"] = f"={ship_ref(y, 'ez', m)}"
    f["W"] = f"=SUM(T{rr}:V{rr})"
    f["X"] = f"=W{rr}/Q{rr}"
    f["Y"] = f"=ROUNDUP(T{rr}/{I['kg_5g']},0)+IF(ISNUMBER(U{rr}),ROUNDUP(U{rr}/{I['kg_5g']},0),0)"
    f["Z"] = f"=Y{rr}*{I['t_5g']}"
    f["AA"] = f"=ROUNDUP(V{rr}/{I['kg_200']},0)"
    f["AB"] = f"=IF(AND(A{rr}>={I['ars_start']},ISNUMBER({I['t_ars']})),{I['t_ars']},{I['t_200']})"
    f["AC"] = f'=IF(A{rr}>={I["ars_start"]},IF(ISNUMBER({I["t_ars"]}),"ARS","ARS (시간 미확인→수동 8 h 참고)"),"수동")'
    f["AD"] = f"=IF({I['sw_200']}=1,AA{rr}*AB{rr},ROUNDUP(V{rr}/Q{rr},0)*AB{rr})"
    f["AE"] = (f'=IF(ISNUMBER(O{rr}),O{rr}*({I["t_iqc"]}/IF(ISNUMBER({I["n_iqc"]}),{I["n_iqc"]},1)+{I["t_fqc"]}+{I["t_oqc"]}),"미확인")')
    f["AF"] = f'=IF(ISNUMBER({I["gb_h"]}),{I["gb_h"]},"미확인")'
    f["AG"] = f'=IF(ISNUMBER(AF{rr}),AF{rr}-Z{rr},"미확인")'
    f["AH"] = f'=IF(ISNUMBER(S{rr}),S{rr}-W{rr},"미확인")'
    f["AI"] = (f'=IF(ISNUMBER(AH{rr}),AH{rr},"보류")' if first else
               f'=IF(AND(ISNUMBER(AH{rr}),ISNUMBER(AI{p})),AI{p}+AH{rr},"보류")')
    f["AJ"] = (f'=IF(ISNUMBER({I["inv0"]}),{I["inv0"]},"미입력")' if first else f'=IF(ISNUMBER(AQ{p}),AQ{p},"보류")')
    f["AK"] = f'=IF(AND(ISNUMBER(AJ{rr}),ISNUMBER(S{rr})),AJ{rr}+S{rr},"보류")'
    f["AL"] = f'=IF(ISNUMBER(AK{rr}),MIN(AK{rr},T{rr}),"보류")'
    f["AM"] = f'=IF(ISNUMBER(AK{rr}),MIN(AK{rr}-AL{rr},N({I["tgt_hx"]})),"보류")'
    f["AN"] = f'=IF(ISNUMBER(AK{rr}),MIN(AK{rr}-AL{rr}-AM{rr},N(U{rr})),"보류")'
    f["AO"] = f'=IF(ISNUMBER(AK{rr}),MIN(AK{rr}-AL{rr}-AM{rr}-AN{rr},V{rr}),"보류")'
    f["AP"] = f'=IF(ISNUMBER(AK{rr}),W{rr}-AL{rr}-AN{rr}-AO{rr},"보류")'
    f["AQ"] = f'=IF(ISNUMBER(AK{rr}),AK{rr}-AL{rr}-AN{rr}-AO{rr},"보류")'
    f["AR"] = f'=IF(ISNUMBER({I["tgt_all"]}),{I["tgt_all"]},"미입력")'
    f["AS"] = f'=IF(AND(ISNUMBER(AJ{rr}),ISNUMBER(AR{rr})),W{rr}+AR{rr}-AJ{rr},"보류")'
    f["AT"] = f'=IF(AND(ISNUMBER(AQ{rr}),ISNUMBER(AR{rr})),AQ{rr}-AR{rr},"보류")'
    f["AU"] = f'=IF(ISNUMBER(AQ{rr}),"계산 가능",IF(ISNUMBER(AH{rr}),"부분 (재고 입력 필요)","보류 (생산 조건 미확인)"))'
    f["AV"] = ('=' + '&'.join([
        f'IF(D{rr}="리플럭스 공사",IF(ISNUMBER({I["h_con"]}),"","공사 중 가용시간·"),"")',
        f'IF(D{rr}="시운전",IF(ISNUMBER({I["b_trial"]}),"","시운전 양품 Batch·"),"")',
        f'IF(D{rr}="리플럭스 적용",IF(ISNUMBER({I["t_ref_rf"]}),"","리플럭스 정제시간·"),"")',
        f'IF(AND(D{rr}="리플럭스 적용",NOT(ISNUMBER({I["kg_b_rf"]}))),"Batch량(190 유지 가정)·","")',
        f'IF(ISNUMBER(N{rr}),"","실적 Batch·")',
        f'IF(ISNUMBER(F{rr}),"","정지·보수(미반영)·")',
        f'IF(ISNUMBER({I["y_all"]}),"","양품률(100% 상한)·")',
        f'IF(ISNUMBER(U{rr}),"","CXMT 물량 미제시·")',
        f'IF(ISNUMBER({I["gb_h"]}),"","글로브 박스 가용시간·")',
        f'IF(ISNUMBER({I["inv0"]}),"","기초재고·")']))
    cols = [CL(i) for i in range(1, len(mh) + 1)]
    nfs = {"A": "ym", "B": "int", "C": "int", "E": "h", "F": "h", "G": "h", "H": "h2", "I": "h", "J": "int", "K": "h", "L": "int",
           "M": "int", "N": "int", "O": "int", "Q": "kg0", "R": "kg0", "S": "kg0", "T": "kg", "U": "kg", "V": "kg", "W": "kg",
           "X": "b", "Y": "int", "Z": "h", "AA": "int", "AB": "h", "AD": "h", "AE": "h", "AF": "h", "AG": "h",
           "AH": "kg", "AI": "kg", "AJ": "kg", "AK": "kg", "AL": "kg", "AM": "kg", "AN": "kg", "AO": "kg", "AP": "kg", "AQ": "kg",
           "AR": "kg", "AS": "kg", "AT": "kg"}
    for col in cols:
        v = f.get(col)
        c = wsM[f"{col}{rr}"]
        c.value = v
        kind = "unk" if col in ("F", "N") else ("in" if col == "A" else ("link" if col in ("T", "U", "V", "L", "M") else "calc"))
        if col in ("J", "R", "S", "W", "AQ"):
            kind = "key"
        style_cell(c, kind, nfs.get(col))
MN = M0 + 23
# summary rows
SR = {2026: MN + 1, 2027: MN + 2}
for yr, rr in SR.items():
    a, b = (M0, M0 + 11) if yr == 2026 else (M0 + 12, MN)
    wsM[f"A{rr}"] = f"{yr} 합계"; style_cell(wsM[f"A{rr}"], "text", bold=True)
    for col in ("J", "O", "R", "S"):
        wsM[f"{col}{rr}"] = (f'=IF(COUNT({col}{a}:{col}{b})=12,SUM({col}{a}:{col}{b}),"부분합계 "&TEXT(SUM({col}{a}:{col}{b}),"#,##0")'
                             f'&" ("&COUNT({col}{a}:{col}{b})&"/12개월)")')
        style_cell(wsM[f"{col}{rr}"], "key", "kg0")
    for col in ("T", "V", "W", "X", "Y", "Z", "AA", "AD") + (("L", "M") if yr == 2026 else ()):
        wsM[f"{col}{rr}"] = f"=SUM({col}{a}:{col}{b})"; style_cell(wsM[f"{col}{rr}"], "key", nfs.get(col))
    wsM[f"U{rr}"] = f'=IF(COUNT(U{a}:U{b})=0,"미제시",SUM(U{a}:U{b}))'; style_cell(wsM[f"U{rr}"], "key", "kg")
# max rows used by 03
MX = {2026: MN + 3, 2027: MN + 4}
for yr, rr in MX.items():
    a, b = (M0, M0 + 11) if yr == 2026 else (M0 + 12, MN)
    wsM[f"A{rr}"] = f"{yr} 월 최대"; style_cell(wsM[f"A{rr}"], "text", bold=True)
    for col in ("W", "Y", "Z", "AA", "AD", "AE"):
        wsM[f"{col}{rr}"] = f"=MAX({col}{a}:{col}{b})"; style_cell(wsM[f"{col}{rr}"], "calc", nfs.get(col))
assert MX[2026] == 32 and MX[2027] == 33, MX
wsM.freeze_panes = "E6"
wsM.auto_filter.ref = f"A5:{CL(len(mh))}{MN}"
for i in range(1, len(mh) + 1):
    wsM.column_dimensions[CL(i)].width = 11
for k, v in {"A": 9, "D": 13, "P": 14, "AC": 16, "AU": 18, "AV": 60}.items():
    wsM.column_dimensions[k].width = v
# conditional formats
red = PatternFill("solid", fgColor="FFC7CE"); redf = Font(color="9C0006")
for col in ("AG", "AH", "AP", "AT"):
    rngc = f"{col}{M0}:{col}{MN}"
    wsM.conditional_formatting.add(rngc, FormulaRule(formula=[f"AND(ISNUMBER({col}{M0}),{col}{M0}<0)"], fill=red, font=redf))
wsM.conditional_formatting.add(f"AP{M0}:AP{MN}", FormulaRule(formula=[f"AND(ISNUMBER(AP{M0}),AP{M0}>0)"], fill=red, font=redf))
grayf = PatternFill("solid", fgColor="EDEDED")
wsM.conditional_formatting.add(f"G{M0}:AU{MN}", FormulaRule(formula=[f'OR(G{M0}="미확인",G{M0}="보류",G{M0}="미입력",LEFT(G{M0},2)="보류")'],
                                                           fill=grayf, font=Font(color="C55A11", italic=True)))
rnote = MX[2027] + 2
for k, t in enumerate([
    "※ 생산 가능 Batch는 '정제기 생산 가능량'이며 실제 생산계획·실적이 아님. 실적은 N열(확정 생산 Batch)에 입력하면 우선 적용.",
    "※ 재고 배정 순서: ① 하이닉스 출하 → ② 하이닉스 목표재고 확보 → ③ CXMT → ④ 이지켐. 기말재고 = 기초재고 + 출하 가능 양품 생산량 − 출하량(배정분).",
    "※ 필요 생산량 = 출하량 + 목표 기말재고 − 기초재고 (목표재고는 재고 수준으로만 반영 — 매년 신규 생산량으로 반복 가산하지 않음).",
    "※ 동일 Batch를 여러 고객에게 배분하므로 필요 Batch는 출하 합계(kg) ÷ Batch량으로 산정 (고객별 올림 합산으로 정제시간을 중복 계산하지 않음).",
    "※ 5 Gal 병수·충진시간은 고객별 20 kg 단위 올림. 200 L 작업시간은 01_Inputs '200 L 8 h 적용 단위' 선택에 따름.",
    "※ 2027년 2~6월(공사·시운전)과 7월 이후(리플럭스 정제시간) 생산 조건은 미확인 → 연간 실제 생산량은 '부분합계'로 표시. 시나리오 비교는 05 시트."]):
    c = wsM.cell(row=rnote + k, column=1, value=t); c.font = font(GRAY, italic=True)

# ======================================================== 05_Reflux_Scenarios
wsR = wb.create_sheet("05_Reflux_Scenarios")
title(wsR, "05_Reflux_Scenarios | 현재 → 공사 · 시운전 → 리플럭스 → 105℃ 검토안",
      "연간 환산 Capa.(47.2톤·50.2톤)와 2027년 달력연도 실제 생산량을 구분. 시나리오 가정은 확정값이 아니며 04 시트 입력란을 대신하지 않음.")
legend(wsR, 3)
widths(wsR, {"A": 34, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 13, "I": 13, "J": 13, "K": 13, "L": 13,
             "M": 13, "N": 13, "O": 14, "P": 40})
row = 5
wsR.cell(row=row, column=1, value="1. 기간 정의").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["구분", "시작", "종료", "운전 조건", "정제기", "충진 (200 L)", "생산 조건 입력", "상태"]); row += 1
periods = [("현재", dt.date(2026, 1, 1), f"={I['con_s']}-1", "103℃ · 57 h/Batch", "공용 1대", "수동 → '27.1~ ARS", "현재 조건 (03 시트 선택 간격)", "제공"),
           ("리플럭스 공사", f"={I['con_s']}", f"={I['con_e']}", "공사 중 가동 여부 미확인", "공용 1대", "ARS", f'=IF(ISNUMBER({I["h_con"]}),"가용 "&{I["h_con"]}&" h/월","공사 중 가용시간 미확인")', "계획"),
           ("시운전", f"={I['trial_s']}", f"={I['rf_s']}-1", "시운전", "공용 1대", "ARS", f'=IF(ISNUMBER({I["b_trial"]}),{I["b_trial"]}&" Batch","시운전 양품 Batch 미확인")', "계획"),
           ("리플럭스 적용 생산", f"={I['rf_s']}", "", "리플럭스 · 103℃", "동일 1대", "ARS", f'=IF(ISNUMBER({I["t_ref_rf"]}),{I["t_ref_rf"]}&" h/Batch","정제시간 미확인")', "계획"),
           ("최초 12개월", f"={I['rf_s']}", f"={I['f12_e']}", "리플럭스", "동일 1대", "ARS", "연간 환산 47.2톤 기준 기간", "제공"),
           ("105℃ 온도 상승 (검토안)", f'=IF(ISNUMBER({I["temp_date"]}),{I["temp_date"]},"미정")', "", "105℃ · 검증·승인 필요", "동일 1대", "ARS", "적용 시점 미정", "검토안")]
for p_ in periods:
    for j, v in enumerate(p_):
        c = wsR.cell(row=row, column=1 + j, value=v)
        style_cell(c, "text" if j not in (1, 2) else "link", "yyyy-mm-dd" if j in (1, 2) else None)
    row += 1

row += 1
wsR.cell(row=row, column=1, value="2. Capa. 로드맵 (연간 환산 기준, 12개월 동일 조건 적용값)").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["구분", "연간 환산\nCapa. (t/년)", "증가분\n(t/년)", "평균 월\n(t/월)", "Batch 상당/년\n(190 kg 유지)", "등가 Batch\n간격 (h, 역산)", "상태", "비고"]); row += 1
RM0 = row
for lab, fx, add, st, note in (("현재 (103℃)", f"={I['capa_now']}", "", "제공 기준값", "과거 일정 관측 기준 연 생산량과 차이 확인 (03 시트 5)"),
                               ("리플럭스 적용 후", f"={I['capa_rf']}", f"={I['capa_add']}", "제공 (12개월 환산)", "'27.7~'28.6 · 2027년 실제 생산량 아님 · ARS 미가산"),
                               ("105℃ 적용 시 (검토안)", f"={I['capa_t']}", f"={I['capa_tadd']}", "검토안 · 미확정", "열 안정성·Dimer·Unknown impurity·수율 검증, 승인 필요, 시점 미정"),
                               ("목표 (2028년 이후)", f"={I['target']}", "", "목표", "50톤 수준")):
    vals = [lab, fx, add, f"=B{row}/12", f"=B{row}*1000/{I['kg_b']}", f"={I['yh']}/E{row}", st, note]
    for j, v in enumerate(vals):
        c = wsR.cell(row=row, column=1 + j, value=v)
        style_cell(c, "text" if j in (0, 6, 7) else ("key" if j == 1 else "calc"), [None, "t", "t", "t", "b", "h2", None, None][j])
    if "검토안" in lab:
        for j in range(8):
            wsR.cell(row=row, column=1 + j).fill = FILL["rev"]
    row += 1
style_cell(wsR.cell(row=row, column=1, value="※ 등가 간격은 190 kg/Batch·8,760 h 연속 가정의 역산 참고값. 리플럭스 후 실제 정제시간·Batch량은 미확인 (01_Inputs D)."), "note")
row += 2

# 3. 2027 actual (link from 04)
wsR.cell(row=row, column=1, value="3. 2027년 달력연도 실제 생산 가능량 (04 시트 입력 기준) — 미확인 월이 있으면 부분합계").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["구분"] + [f"{m}월" for m in range(1, 13)] + ["연간", "비고"]); row += 1
A27 = row
rows27 = [("운전 조건", "D", None), ("생산 가능 Batch", "O", "int"), ("양품 생산량 (kg)", "S", "kg0"), ("출하 합계 (kg)", "W", "kg0")]
for lab, col, nf in rows27:
    style_cell(wsR.cell(row=row, column=1, value=lab), "text", bold=True)
    for m in range(12):
        c = wsR.cell(row=row, column=2 + m, value=f"='04_Monthly_2026_2027'!{col}{M0+12+m}"); style_cell(c, "link", nf)
    if col in ("O", "S"):
        c = wsR.cell(row=row, column=14, value=f"='04_Monthly_2026_2027'!{'O' if col=='O' else 'S'}{SR[2027]}")
    elif col == "W":
        c = wsR.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    else:
        c = wsR.cell(row=row, column=14, value="-")
    style_cell(c, "key", nf)
    row += 1
style_cell(wsR.cell(row=row, column=1, value="※ 공사·시운전 월 생산량은 0 또는 정상값으로 임의 입력하지 않음 (01_Inputs E 입력 시 자동 반영)."), "note")
row += 2

# 4. scenarios
wsR.cell(row=row, column=1, value="4. 2027년 공급 대응 시나리오 (참고 · 확정 아님) — 필요 선행재고 = 누적 (생산 − 출하)의 최저점").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["시나리오 가정", "S1 공사 중 중단\n7월 즉시 47.2 환산", "S2 공사 중 현재 유지\n7월 즉시 47.2 환산", "S3 공사 중 50%\n안정화 3개월", "", "", "", "", "", "", "", "", "", "", "", "설명"]); row += 1
SP = {}
params = [("sp_con", "공사 기간 가동률 (현재 조건 대비)", [0, 1, 0.5], "pct", "현재 조건 월 생산 가능량 × 가동률 (2~5월)"),
          ("sp_trial", "시운전 월 양품 생산 (kg)", [0, 0, 0], "kg0", "6월 — 양품 출하 가능 여부 미확인 → 0 가정"),
          ("sp_stab", "리플럭스 초기 안정화 개월 (현재 수준 생산)", [0, 0, 3], "int", "7월부터 해당 개월 동안 현재 조건 생산량 적용"),
          ("sp_rfkg", "리플럭스 안정 후 월 생산 (kg/월)", [f"={I['capa_rf']}*1000/12"] * 3, "kg0", "47.2톤 연간 환산 ÷ 12 — 상한 참고"),
          ("sp_y", "양품률", [1, 1, 1], "pct", "100% = 상한")]
for key, lab, vals, nf, desc in params:
    style_cell(wsR.cell(row=row, column=1, value=lab), "text", bold=True)
    for j, v in enumerate(vals):
        c = wsR.cell(row=row, column=2 + j, value=v); style_cell(c, "est", nf)
    wsR.merge_cells(start_row=row, start_column=5, end_row=row, end_column=16)
    style_cell(wsR.cell(row=row, column=5, value=desc), "note")
    SP[key] = row
    row += 1
row += 1
# base current-condition monthly capacity (kg) for 2027 months (decimal batch equivalent)
header(wsR, row, ["2027 월별 (kg)"] + [f"{m}월" for m in range(1, 13)] + ["연간", "비고"]); row += 1
BASE = row
style_cell(wsR.cell(row=row, column=1, value="현재 조건 월 생산 가능량 (Batch 상당 × 190)"), "text", bold=True)
for m in range(12):
    c = wsR.cell(row=row, column=2 + m, value=f"='04_Monthly_2026_2027'!C{M0+12+m}*24/{C['basis_int']}*{I['kg_b']}"); style_cell(c, "calc", "kg0")
c = wsR.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})"); style_cell(c, "calc", "kg0")
style_cell(wsR.cell(row=row, column=15, value="03 시트 선택 간격 (월 이월 평균화)"), "note")
row += 1
SHIPR = row
style_cell(wsR.cell(row=row, column=1, value="출하 합계 (하이닉스+CXMT+이지켐)"), "text", bold=True)
for m in range(12):
    c = wsR.cell(row=row, column=2 + m, value=f"='01_Inputs'!{CL(3+m)}{SHIP[(2027,'tot')]}"); style_cell(c, "link", "kg0")
c = wsR.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})"); style_cell(c, "key", "kg0")
row += 1
SCN = {}
for j, sname in enumerate(["S1", "S2", "S3"]):
    pc = CL(2 + j)
    # production row
    pr = row
    style_cell(wsR.cell(row=row, column=1, value=f"{sname} 생산 (kg)"), "text", bold=True)
    for m in range(12):
        col = CL(2 + m)
        mm = m + 1
        fx = (f"=IF({mm}=1,{col}{BASE},IF({mm}<=5,{col}{BASE}*${pc}${SP['sp_con']},IF({mm}=6,${pc}${SP['sp_trial']},"
              f"IF({mm}-6<=${pc}${SP['sp_stab']},{col}{BASE},${pc}${SP['sp_rfkg']}))))*${pc}${SP['sp_y']}")
        c = wsR.cell(row=row, column=2 + m, value=fx); style_cell(c, "calc", "kg0")
    c = wsR.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})"); style_cell(c, "key", "kg0")
    row += 1
    cr = row
    style_cell(wsR.cell(row=row, column=1, value=f"{sname} 누적 (생산 − 출하)"), "text")
    for m in range(12):
        col = CL(2 + m)
        fx = f"={col}{pr}-{col}{SHIPR}" if m == 0 else f"={CL(1+m)}{cr}+{col}{pr}-{col}{SHIPR}"
        c = wsR.cell(row=row, column=2 + m, value=fx); style_cell(c, "calc", "kg0")
    row += 1
    SCN[sname] = (pr, cr)
wsR.conditional_formatting.add(f"B{SCN['S1'][1]}:M{SCN['S3'][1]}", CellIsRule(operator="lessThan", formula=["0"], fill=red, font=redf))
row += 1
header(wsR, row, ["시나리오 결과", "S1 공사 중 중단\n7월 즉시 47.2 환산", "S2 공사 중 현재 유지\n7월 즉시 47.2 환산", "S3 공사 중 50%\n안정화 3개월", "", "", "", "", "", "", "", "", "", "", "", "해석"]); row += 1
RES = {}
res_rows = [("2027 생산 (kg, 참고)", lambda s: f"=N{SCN[s][0]}", "kg0", "시나리오 가정 기준 — 실제 생산량 아님"),
            ("2027 출하계획 (kg)", lambda s: f"=N{SHIPR}", "kg0", "38,020 kg"),
            ("연간 생산 − 출하 (kg)", lambda s: f"=N{SCN[s][0]}-N{SHIPR}", "kg0", "음수 = 연간 부족"),
            ("필요 선행재고 (2027-01-01 기준, kg)", lambda s: f"=MAX(0,-MIN(B{SCN[s][1]}:M{SCN[s][1]}))", "kg0", "월별 결품 없이 대응하기 위한 최소 기초재고"),
            ("필요 선행재고 (190 kg Batch 상당)", lambda s: f"=B{{prev}}/{I['kg_b']}", "b", ""),
            ("최대 부족 발생 월", lambda s: f'=IF(MIN(B{SCN[s][1]}:M{SCN[s][1]})<0,MATCH(MIN(B{SCN[s][1]}:M{SCN[s][1]}),B{SCN[s][1]}:M{SCN[s][1]},0)&"월","-")', None, "누적 최저점")]
for lab, fn, nf, desc in res_rows:
    style_cell(wsR.cell(row=row, column=1, value=lab), "text", bold=True)
    for j, s in enumerate(["S1", "S2", "S3"]):
        fx = fn(s).replace("B{prev}", f"{CL(2+j)}{row-1}")
        c = wsR.cell(row=row, column=2 + j, value=fx); style_cell(c, "key", nf)
    wsR.merge_cells(start_row=row, start_column=5, end_row=row, end_column=16)
    style_cell(wsR.cell(row=row, column=5, value=desc), "note")
    RES[lab] = row
    row += 1
# H2 requirement
H2 = {}
r0 = row
h2rows = [("2027 7~12월 출하 합계 (kg)", f"=SUM(H{SHIPR}:M{SHIPR})", "kg0", "리플럭스 적용 기간 출하"),
          ("7~12월 달력시간 (h)", f"=SUM('04_Monthly_2026_2027'!E{M0+18}:E{MN})", "int", "184일 × 24 h"),
          ("7~12월 필요 Batch (190 kg · 양품률 100% · 재고 미사용)", f"=B{r0}/{I['kg_b']}", "b", "Batch량 변경 시 01_Inputs 반영"),
          ("7~12월 월 평균 필요 Batch", f"=B{r0+2}/6", "b", ""),
          ("출하 대응 필요 평균 Batch 간격 (h/Batch)", f"=B{r0+1}/B{r0+2}", "h2", "리플럭스 후 실효 간격이 이 값 이하여야 당월 생산으로 출하 대응 (현재 관측 약 67.5 h)")]
for lab, fx, nf, desc in h2rows:
    style_cell(wsR.cell(row=row, column=1, value=lab), "text", bold=True)
    c = wsR.cell(row=row, column=2, value=fx); style_cell(c, "key", nf)
    wsR.merge_cells(start_row=row, start_column=3, end_row=row, end_column=16)
    style_cell(wsR.cell(row=row, column=3, value=desc), "note")
    H2[lab] = row
    row += 1
row += 1
wsR.cell(row=row, column=1, value="5. 105℃ 온도 상승 검토안 — 검증 필요사항 (미확정 · 검증 및 승인 필요 · 적용 시점 미정)").font = font("EB002C", True, 11); row += 1
for t in ["열 안정성 (103→105℃ 가열 시 분해·변색)", "Dimer 형성 증가 여부", "Unknown impurity 증가 여부 (NMR·ICP-MS)",
          "수율 영향 및 Batch량 변화", "고객 승인 (하이닉스·CXMT·이지켐 품질 기준)", "적용 시점 — 미정 (2028년 검토)"]:
    style_cell(wsR.cell(row=row, column=1, value="· " + t), "text"); row += 1

# ======================================================== 06_Report_Summary
wsS = wb.create_sheet("06_Report_Summary")
title(wsS, "06_Report_Summary | PPT 사용 수치 · 차트 (다른 시트 참조 수식)", "PPT 본문 수치는 이 시트 값과 일치해야 함.")
widths(wsS, {"A": 46, "B": 14, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12, "M": 12, "N": 13, "O": 30})
row = 4
wsS.cell(row=row, column=1, value="1. 핵심 수치 (KPI)").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["항목", "값", "단위", "출처 시트"]); row += 1
KPI = {}
kpis = [("현재 연간 Capa.", f"={I['capa_now']}", "t/년", "01", "t"),
        ("26.2톤 Batch 상당", f"={C['cp_b']}", "Batch/년", "03", "0.0"),
        ("26.2톤 반올림 정수 Batch 생산량", f"={C['cp_kground']}", "kg/년", "03", "kg0"),
        ("26.2톤 역산 간격", f"={C['cp_int']}", "h/Batch", "03", "h2"),
        ("월 평균 Batch 상당 (26.2톤)", f"={C['cp_bm']}", "Batch/월", "03", "b"),
        ("5 Gal 경로 확인 Lead Time", f"={C['lt_5g']}", "h", "03", "h"),
        ("200 L 경로 확인 Lead Time", f"={C['lt_200']}", "h", "03", "h"),
        ("Lead Time 미확인 항목 수", f"={C['lt_unk']}", "개", "03", "int"),
        ("이론 Batch/월 (소수)", f"={C['th_bm']}", "Batch/월", "03", "b"),
        ("이론 월 정수 Batch", f"={C['th_bi']}", "Batch/월", "03", "int"),
        ("이론 월말 이월 h", f"={C['th_carry_h']}", "h", "03", "h"),
        ("이론 월 생산량", f"={C['th_kgm']}", "kg/월", "03", "kg0"),
        ("이론 연 Batch (달력)", f"={C['th_by']}", "Batch/년", "03", "int"),
        ("이론 연 생산량 (달력)", f"={C['th_kgy']}", "kg/년", "03", "kg0"),
        ("이론 단순 12배 생산량", f"={C['th_kgy12']}", "kg/년", "03", "kg0"),
        ("11 Batch 정제시간", f"={C['mx_ref']}", "h", "03", "h"),
        ("11 Batch 720 h 차이", f"={C['mx_gap']}", "h", "03", "h"),
        ("11 Batch 실효시간", f"={C['mx_int']}", "h/Batch", "03", "h2"),
        ("11 Batch 실효-정제 차이", f"={C['mx_diff']}", "h/Batch", "03", "h2"),
        ("11 Batch 월 생산량", f"={C['mx_kgm']}", "kg/월", "03", "kg0"),
        ("11 Batch 단순 12배", f"={C['mx_kgy12']}", "kg/년", "03", "kg0"),
        ("11 Batch 달력 연속", f"={C['mx_kgy']}", "kg/년", "03", "kg0"),
        ("관측 추정 간격", f"={C['ob_int']}", "h/Batch", "03", "h2"),
        ("관측 추정 하한", f"={C['ob_lo']}", "h/Batch", "03", "h2"),
        ("관측 추정 상한", f"={C['ob_hi']}", "h/Batch", "03", "h2"),
        ("관측 경과일수", f"={C['ob_days']}", "일", "03", "int"),
        ("관측 간격 수", f"={C['ob_n']}", "간격", "03", "int"),
        ("관측 비정제 시간", f"={C['ob_ovh']}", "h/Batch", "03", "h2"),
        ("관측 월 Batch 30일", f"={C['ob_b30']}", "Batch/월", "03", "b"),
        ("관측 월 Batch 31일", f"={C['ob_b31']}", "Batch/월", "03", "b"),
        ("관측 연 Batch", f"={C['ob_by']}", "Batch/년", "03", "int"),
        ("관측 연 생산량", f"={C['ob_kgy']}", "kg/년", "03", "kg0"),
        ("관측 3일 간격 비중", f"={C['ob_share3']}", "%", "03", "pct"),
        ("관측 중앙값 h", f"={C['ob_med']}", "h", "03", "h"),
        ("#37 포함 평균", f"={C['ob_alt37']}", "h", "03", "h2"),
        ("Total 기준 간격", f"={C['ob_tot_int']}", "h", "03", "h2"),
        ("57 h 정합 확인 구간 수", f"={C['ob_dense']}", "구간", "02", "int"),
        ("대조 차이 건수", f"={CHK[0]}", "건", "02", "int"),
        ("47.2톤 등가 간격", f"={C['rf_eq_int']}", "h/Batch", "03", "h2"),
        ("Tank 1교대 여유 vs 57", f"='03_Capacity_Model'!$E${C['tk0']}", "h", "03", "h"),
        ("Tank 1교대 여유 vs 관측", f"='03_Capacity_Model'!$F${C['tk0']}", "h", "03", "h"),
        ("Tank 1교대 여유 vs 47.2", f"='03_Capacity_Model'!$G${C['tk0']}", "h", "03", "h"),
        ("Tank 2교대 여유 vs 47.2", f"='03_Capacity_Model'!$G${C['tk0']+1}", "h", "03", "h"),
        ("Tank 점유 확인분", f"={C['tk_occ']}", "h", "03", "h"),
        ("2026 5 Gal 월 최대 작업", f"={C['pp_gb_need26']}", "h/월", "04", "h"),
        ("2027 5 Gal 월 최대 작업", f"={C['pp_gb_need27']}", "h/월", "04", "h"),
        ("2027 200 L 월 최대 작업", f"={C['pp_200_need27']}", "h/월", "04", "h"),
        ("1교대 가용 h", f"='03_Capacity_Model'!$D${C['pp0']+1}", "h/월", "03", "h"),
        ("2교대 가용 h", f"='03_Capacity_Model'!$D${C['pp0']+2}", "h/월", "03", "h"),
        ("2027 5 Gal 부하율 1교대", f"='03_Capacity_Model'!$F${C['pp0']+1}", "%", "03", "pct"),
        ("2027 5 Gal 부하율 2교대", f"='03_Capacity_Model'!$F${C['pp0']+2}", "%", "03", "pct"),
        ("2026 5 Gal 부하율 1교대", f"='03_Capacity_Model'!$E${C['pp0']+1}", "%", "03", "pct"),
        ("2026 제시 출하", f"='01_Inputs'!$O${SHIP[(2026,'tot')]}", "kg", "01", "kg0"),
        ("2027 출하 합계", f"='01_Inputs'!$O${SHIP[(2027,'tot')]}", "kg", "01", "kg0"),
        ("2027 하이닉스", f"='01_Inputs'!$O${SHIP[(2027,'hx')]}", "kg", "01", "kg0"),
        ("2027 CXMT", f"='01_Inputs'!$O${SHIP[(2027,'cx')]}", "kg", "01", "kg0"),
        ("2027 이지켐", f"='01_Inputs'!$O${SHIP[(2027,'ez')]}", "kg", "01", "kg0"),
        ("2026 생산 가능량 합계 (선택 기준)", f"='04_Monthly_2026_2027'!$S${SR[2026]}", "kg", "04", "kg0"),
        ("2026 생산 가능 Batch 합계", f"='04_Monthly_2026_2027'!$O${SR[2026]}", "Batch", "04", "int"),
        ("2027 생산 가능량 (부분합계)", f"='04_Monthly_2026_2027'!$S${SR[2027]}", "kg", "04", None),
        ("리플럭스 후 Capa.", f"={I['capa_rf']}", "t/년", "01", "t"),
        ("105℃ Capa.", f"={I['capa_t']}", "t/년", "01", "t"),
        ("47.2 월 환산", f"={I['capa_rf']}/12", "t/월", "05", "t"),
        ("47.2 Batch 상당/월", f"={I['capa_rf']}*1000/12/{I['kg_b']}", "Batch/월", "05", "b"),
        ("S1 2027 생산", f"=B{RES['2027 생산 (kg, 참고)']}".replace("B", "'05_Reflux_Scenarios'!B", 1), "kg", "05", "kg0"),
        ("S2 2027 생산", f"='05_Reflux_Scenarios'!C{RES['2027 생산 (kg, 참고)']}", "kg", "05", "kg0"),
        ("S3 2027 생산", f"='05_Reflux_Scenarios'!D{RES['2027 생산 (kg, 참고)']}", "kg", "05", "kg0"),
        ("S1 연간 차이", f"='05_Reflux_Scenarios'!B{RES['연간 생산 − 출하 (kg)']}", "kg", "05", "kg0"),
        ("S2 연간 차이", f"='05_Reflux_Scenarios'!C{RES['연간 생산 − 출하 (kg)']}", "kg", "05", "kg0"),
        ("S3 연간 차이", f"='05_Reflux_Scenarios'!D{RES['연간 생산 − 출하 (kg)']}", "kg", "05", "kg0"),
        ("S1 필요 선행재고", f"='05_Reflux_Scenarios'!B{RES['필요 선행재고 (2027-01-01 기준, kg)']}", "kg", "05", "kg0"),
        ("S2 필요 선행재고", f"='05_Reflux_Scenarios'!C{RES['필요 선행재고 (2027-01-01 기준, kg)']}", "kg", "05", "kg0"),
        ("S3 필요 선행재고", f"='05_Reflux_Scenarios'!D{RES['필요 선행재고 (2027-01-01 기준, kg)']}", "kg", "05", "kg0"),
        ("S1 최대 부족 월", f"='05_Reflux_Scenarios'!B{RES['최대 부족 발생 월']}", "", "05", None),
        ("S2 최대 부족 월", f"='05_Reflux_Scenarios'!C{RES['최대 부족 발생 월']}", "", "05", None),
        ("S3 최대 부족 월", f"='05_Reflux_Scenarios'!D{RES['최대 부족 발생 월']}", "", "05", None),
        ("2027 1월 생산 가능 Batch", f"='04_Monthly_2026_2027'!$O${M0+12}", "Batch", "04", "int"),
        ("2027 1월 생산 가능량", f"='04_Monthly_2026_2027'!$S${M0+12}", "kg", "04", "kg0"),
        ("현재 조건 적용 간격", f"={C['basis_int']}", "h/Batch", "03", "h2"),
        ("2027 H1 출하", f"=SUM('01_Inputs'!C{SHIP[(2027,'tot')]}:H{SHIP[(2027,'tot')]})", "kg", "01", "kg0"),
        ("2027 H2 출하", f"=SUM('01_Inputs'!I{SHIP[(2027,'tot')]}:N{SHIP[(2027,'tot')]})", "kg", "01", "kg0"),
        ("7~12월 월 평균 필요 Batch", f"='05_Reflux_Scenarios'!B{H2['7~12월 월 평균 필요 Batch']}", "Batch/월", "05", "b"),
        ("7~12월 필요 간격", f"='05_Reflux_Scenarios'!B{H2['출하 대응 필요 평균 Batch 간격 (h/Batch)']}", "h/Batch", "05", "h2"),
        ("7~12월 필요 Batch 합계", f"='05_Reflux_Scenarios'!B{H2['7~12월 필요 Batch (190 kg · 양품률 100% · 재고 미사용)']}", "Batch", "05", "b"),
        ("2027 5 Gal 월 병수 (4월~)", f"='04_Monthly_2026_2027'!$Y${M0+12+3}", "병/월", "04", "int"),
        ("2027 200 L 월 용기", f"='04_Monthly_2026_2027'!$AA${M0+12}", "용기/월", "04", "int"),
        ("2026 5 Gal 월 최대 병수", f"='04_Monthly_2026_2027'!$Y${MX[2026]}", "병/월", "04", "int"),
        ]
for lab, fx, unit, src, nf in kpis:
    style_cell(wsS.cell(row=row, column=1, value=lab), "text")
    c = wsS.cell(row=row, column=2, value=fx); style_cell(c, "link", nf)
    style_cell(wsS.cell(row=row, column=3, value=unit), "text")
    style_cell(wsS.cell(row=row, column=4, value=src), "text")
    KPI[lab] = row
    row += 1
row += 1
# 2. capacity comparison (chart source)
wsS.cell(row=row, column=1, value="2. 이론 · 과거 · 추정 생산능력 비교 (연 생산량 t, 달력 연속)").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["구분", "Batch 간격 (h)", "월 Batch (30일)", "연 생산량 (t)", "26.2톤 대비 (t)"]); row += 1
CMPS = row
for k in range(5):
    cr_ = C["cmp0"] + k
    vals = [f"='03_Capacity_Model'!A{cr_}", f"='03_Capacity_Model'!B{cr_}", f"='03_Capacity_Model'!C{cr_}", f"='03_Capacity_Model'!F{cr_}", f"='03_Capacity_Model'!G{cr_}"]
    for j, v in enumerate(vals):
        c = wsS.cell(row=row, column=1 + j, value=v); style_cell(c, "link", [None, "h2", "b", "t", "t"][j])
    row += 1
ch = BarChart(); ch.type = "bar"; ch.title = "연 생산량 비교 (t/년, 190 kg/Batch)"; ch.style = 10
ch.add_data(Reference(wsS, min_col=4, min_row=CMPS - 1, max_row=CMPS + 4), titles_from_data=True)
ch.set_categories(Reference(wsS, min_col=1, min_row=CMPS, max_row=CMPS + 4)); ch.height = 7; ch.width = 16; ch.legend = None
wsS.add_chart(ch, f"G{CMPS-1}")
row += 8
# 3. monthly table 24 months
wsS.cell(row=row, column=1, value="3. 월별 생산 가능량 · 출하 · 재고 증감 (kg)").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["연월", "운전 조건", "생산 가능\nBatch", "양품 생산\n가능량", "하이닉스", "CXMT", "이지켐", "출하 합계", "생산−출하",
                  "누적 증감\n(기초재고 제외)", "과거계획\n번호 개수", "과거계획\nTotal", "기말재고", "계산 가능 여부"]); row += 1
MS0 = row
for k in range(24):
    rr = M0 + k
    src_cols = ["A", "D", "O", "S", "T", "U", "V", "W", "AH", "AI", "L", "M", "AQ", "AU"]
    for j, col in enumerate(src_cols):
        c = wsS.cell(row=row, column=1 + j, value=f"='04_Monthly_2026_2027'!{col}{rr}")
        style_cell(c, "link", ["ym", None, "int", "kg0", "kg", "kg", "kg", "kg", "kg0", "kg0", "int", "int", "kg0", None][j])
    row += 1
wsS.conditional_formatting.add(f"I{MS0}:J{MS0+23}", CellIsRule(operator="lessThan", formula=["0"], fill=red, font=redf))
ch2 = BarChart(); ch2.title = "2026~2027 월별 출하 합계 vs 생산 가능량 (kg)"; ch2.style = 10
ch2.add_data(Reference(wsS, min_col=4, min_row=MS0 - 1, max_row=MS0 + 23), titles_from_data=True)
ch2.add_data(Reference(wsS, min_col=8, min_row=MS0 - 1, max_row=MS0 + 23), titles_from_data=True)
ch2.set_categories(Reference(wsS, min_col=1, min_row=MS0, max_row=MS0 + 23)); ch2.height = 8; ch2.width = 26
wsS.add_chart(ch2, f"P{MS0}")
row += 1
# 4. gap distribution
wsS.cell(row=row, column=1, value="4. 과거 Batch 표시일자 간격 분포 (#39~#92)").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["간격 (일)", "건수", "명목 h", "가능 경과 범위 (h)"]); row += 1
GD0 = row
for j, (lab, col) in enumerate((("1일", "O"), ("2일", "P"), ("3일", "Q"), ("4일", "R"), ("5일", "S"), ("6일 이상", "T"))):
    vals = [lab, f"='02_Batch_Raw'!{col}{SUMROW['tot']}", (j + 1) * 24 if j < 5 else "≥144",
            f"{max(0,j)*24}~{(j+2)*24}" if j < 5 else "≥120"]
    for jj, v in enumerate(vals):
        c = wsS.cell(row=row, column=1 + jj, value=v); style_cell(c, "link" if jj == 1 else "text", "int" if jj in (1, 2) else None)
    row += 1
ch3 = BarChart(); ch3.title = "Batch 표시일자 간격 분포 (건수)"; ch3.style = 10
ch3.add_data(Reference(wsS, min_col=2, min_row=GD0 - 1, max_row=GD0 + 5), titles_from_data=True)
ch3.set_categories(Reference(wsS, min_col=1, min_row=GD0, max_row=GD0 + 5)); ch3.height = 6; ch3.width = 12; ch3.legend = None
wsS.add_chart(ch3, f"G{GD0-1}")
row += 3
# 5. roadmap
wsS.cell(row=row, column=1, value="5. 투자 일정 · Capa. 로드맵").font = font("EB002C", True, 11); row += 1
for lab, fx, nf in (("리플럭스 공사", f"=TEXT({I['con_s']},\"yyyy.m\")&\" ~ \"&TEXT({I['con_e']},\"yyyy.m\")", None),
                    ("시운전", f"=TEXT({I['trial_s']},\"yyyy.m\")", None),
                    ("리플럭스 적용 생산 시작", f"=TEXT({I['rf_s']},\"yyyy.m\")", None),
                    ("최초 12개월", f"=TEXT({I['rf_s']},\"yyyy.m\")&\" ~ \"&TEXT({I['f12_e']},\"yyyy.m\")", None),
                    ("ARS 운영 시작 (계획)", f"=TEXT({I['ars_start']},\"yyyy.m\")", None),
                    ("Capa. 26.2 → 47.2 → (검토안) 50.2 t/년", f"={I['capa_now']}&\" → \"&{I['capa_rf']}&\" → (\"&{I['capa_t']}&\")\"", None)):
    style_cell(wsS.cell(row=row, column=1, value=lab), "text")
    c = wsS.cell(row=row, column=2, value=fx); style_cell(c, "link", nf)
    wsS.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    row += 1
row += 1
# 6. open items
wsS.cell(row=row, column=1, value="6. 주요 미확인 사항 (담당)").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["확인사항", "담당", "영향"]); row += 1
opens = [("Batch 표시일자 의미 (투입/정제 시작/완료/출하) · 실제 시각 · 계획 대비 실적", "생산", "실효 간격 67.5 h 추정의 전제"),
         ("원자료 Total(46) vs 번호 개수(54) 집계 기준 · 색상 의미 · 원본 이미지 대조", "생산", "완료·양품 Batch 수 확정"),
         ("정제 57 h 포함 범위 (준비·투입·가열·냉각·배출) · 세척·전환 시간", "생산", "비정제 시간 약 10.5 h 구성"),
         ("190 kg = 정제 회수량/양품량 여부 · 8병(160 kg)과 30 kg 차이 처리", "생산·품질", "출하 가능량"),
         ("IQC 적용 단위 · 200 L 8 h 적용 단위 · OQC 적용 단위", "품질·생산", "후공정 작업량"),
         ("글로브 박스 수·교대 · Product Tank 수·용량 · 검사 가용시간", "생산", "후공정 병목 (2027 5 Gal 약 248 h/월)"),
         ("공사 중 정제기 가동 여부 · 시운전 양품 · 안정화 기간", "설비·공무", "2027 생산량 (부족 규모 좌우)"),
         ("리플럭스 후 정제시간 · Batch량", "설비·생산", "47.2톤 달성 조건 (등가 약 35 h/Batch)"),
         ("기초재고 · 목표재고 · 고객별 합격률(이지켐 색도)", "생산관리·품질", "월별 재고·부족량"),
         ("2026 CXMT 출하량 · 하이닉스 실적/계획 구분", "영업", "2026 출하 합계"),
         ("ARS 운영 개시일 · 충진시간 (Capa. 미가산)", "설비", "200 L 작업시간"),
         ("105℃: 열 안정성 · Dimer · Unknown impurity · 수율 · 승인 · 적용 시점", "품질·기술", "50.2톤 검토안")]
for t, who, eff in opens:
    for j, v in enumerate((t, who, eff)):
        style_cell(wsS.cell(row=row, column=1 + j, value=v), "text")
    row += 1
wsS.freeze_panes = "A4"

for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = {"01_Inputs": "2E75B6", "02_Batch_Raw": "7F7F7F", "03_Capacity_Model": "EB002C",
                                    "04_Monthly_2026_2027": "FF7900", "05_Reflux_Scenarios": "70AD47", "06_Report_Summary": "404040"}[ws.title]
wb.save(OUT)
import json
json.dump({"KPI": KPI, "M0": M0, "SR": SR, "SUMROW": SUMROW, "SHIP": {f"{k}": v for k, v in SHIP.items()},
           "CMP0": C["cmp0"], "MS0": MS0, "GD0": GD0, "RES": RES, "G0": C["g0"], "GN": C["gn"]},
          open(OUT + ".map.json", "w"), ensure_ascii=False)
print("saved", OUT)
