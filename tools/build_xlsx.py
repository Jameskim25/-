# -*- coding: utf-8 -*-
"""SKTC CpZr 생산능력 계산 파일 (v2: 정제 45 h / 충진 전 57 h, 정제기 2대 시차 운전, 4개 고객, 2026~2028 36개월).

입력값(사용자 제공 조건·생산팀 면담·원자료)을 01_Inputs·02_Batch_Raw에 두고 모든 결과는 Excel 수식으로 계산한다.
사용: python build_xlsx.py <원본1 pptx(분포도)> <원본2 pptx(노트)> <out.xlsx>
"""
import datetime as dt
import json
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.formula import ArrayFormula
from pptx import Presentation

from ppt_sheets import PIN, P3_GRID, P3_GRID_FIRST, P3_TOTAL, add_ppt_sheets, pin_ref, ship_cell

SRC1, SRC2, OUT = sys.argv[1:4]
SRC_NAME = "SKTC_CpZr_Capacity_Reflux_Roadmap_2027_2028.pptx (첨부 원본·수정본)"
USER = "사용자 제공 조건 (2026-10-01 지시)"
USER57 = "사용자가 정한 자료 작성 기준 / 57시간에서 기타 공정시간 차감"
ORAL = "생산팀 면담 구두 언급 (미검증)"

FONT = "Arial"


def font(color="000000", bold=False, italic=False, size=9):
    return Font(name=FONT, color=color, bold=bold, italic=italic, size=size)


FILL = {"in": PatternFill("solid", fgColor="DDEBF7"), "unk": PatternFill("solid", fgColor="FFF2CC"),
        "est": PatternFill("solid", fgColor="FCE4D6"), "rev": PatternFill("solid", fgColor="E4DFEC"),
        "oral": PatternFill("solid", fgColor="E2EFDA"), "hdr": PatternFill("solid", fgColor="404040"),
        "hdr2": PatternFill("solid", fgColor="C55A11"), "sec": PatternFill("solid", fgColor="F2F2F2"),
        "key": PatternFill("solid", fgColor="FDE9E7")}
BLUE, GREEN, BLACK, GRAY = "0000FF", "008000", "000000", "7F7F7F"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
NF = {"kg": '#,##0.0;[Red]-#,##0.0;"0"', "kg0": '#,##0;[Red]-#,##0;"0"', "h": '0.0;[Red]-0.0', "h2": '0.00',
      "b": '0.00', "b1": '0.0', "t": '0.00', "pct": '0.0%', "d": 'yyyy-mm-dd', "ym": 'yyyy-mm', "md": 'm"/"d', "int": '0'}
red = PatternFill("solid", fgColor="FFC7CE"); redf = Font(color="9C0006", bold=True)

wb = Workbook()


def style_cell(c, kind="calc", nf=None, bold=False):
    c.border = BORDER; c.alignment = WRAP
    k = {"in": ("in", BLUE), "unk": ("unk", BLUE), "est": ("est", BLUE), "rev": ("rev", BLUE), "oral": ("oral", BLUE)}
    if kind in k:
        c.fill = FILL[k[kind][0]]; c.font = font(k[kind][1], bold)
    elif kind == "link":
        c.font = font(GREEN, bold)
    elif kind == "key":
        c.font = font(BLACK, True); c.fill = FILL["key"]
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
    items = [("in", "입력 (제공·확정·계획)"), ("unk", "미확인 입력란 (공란 유지)"), ("est", "추정·가정 입력"),
             ("oral", "구두 언급 (미검증)"), ("rev", "검토안 (미확정)"), ("calc", "수식 / 시트 참조(초록)"), ("key", "핵심 결과")]
    for i, (k, lab) in enumerate(items):
        style_cell(ws.cell(row=row, column=col + 2 * i, value="  "), k)
        ws.cell(row=row, column=col + 2 * i + 1, value=lab).font = font(GRAY, size=8)


def widths(ws, spec):
    for k, v in spec.items():
        ws.column_dimensions[k].width = v


# ======================================================== 01_Inputs
wsI = wb.active; wsI.title = "01_Inputs"
title(wsI, "01_Inputs | CpZr 생산능력 계산 입력값 (v2)",
      "정제 45 h / 충진 전까지 57 h · 현재 정제기 1대 → 개선 후 2대 시차 운전 · 4개 고객 · 2026~2028. 미확인은 노란 공란 — 확인 후 입력.")
legend(wsI, 3)
header(wsI, 5, ["항목", "값", "단위", "적용 범위", "구분", "출처", "비고"])
REF = {}
r = 6


def sec(label):
    global r
    for col in range(1, 8):
        c = wsI.cell(row=r, column=col, value=label if col == 1 else None)
        c.fill = FILL["sec"]; c.font = font(BLACK, True); c.border = BORDER
    r += 1


KIND = {"확정": "in", "제공": "in", "계획": "in", "추정": "est", "가정": "est", "미확인": "unk", "검토안": "rev",
        "수식": "calc", "구두": "oral", "연결": "link"}


def inp(key, item, val, unit, scope, kind, src, note="", nf=None):
    global r
    if key in PIN:
        sheet, cell, _, pnf = PIN[key]
        ref = pin_ref(key)
        val = f'=IF(ISBLANK({ref}),"",{ref})'
        kind = "연결"; src = f"{sheet} {cell} 입력"
        note = f"수정은 {sheet} {cell}에서" + (f" · {note}" if note else "")
        if pnf and pnf.startswith("yyyy"):
            nf = "d"
    vals = [item, val, unit, scope, {"구두": "구두 언급"}.get(kind, kind), src, note]
    for j, v in enumerate(vals):
        c = wsI.cell(row=r, column=1 + j, value=v)
        style_cell(c, KIND[kind] if j == 1 else "text", nf if j == 1 else None)
    if isinstance(val, (dt.date, dt.datetime)):
        wsI.cell(row=r, column=2).number_format = NF["d"]
    REF[key] = f"'01_Inputs'!$B${r}"
    r += 1


def R(key):
    return REF[key]


sec("A. 설비 · 가동시간")
inp("n_now", "정제기 수 (As-is)", 1, "대", "~'27.6 · 4개 고객 공동", "제공", USER, "단일 설비 순차 생산", "int")
inp("n_new", "정제기 수 (To-be)", 2, "대", "'27.7~ · 시간차 병행 운전", "제공", USER, "고객별 전용 설비 아님 — 전체 자원 안에서 배정", "int")
inp("temp_now", "현재 정제온도", 103, "℃", "", "제공", USER, "", "int")
inp("mdays", "월 비교 기준 일수", 30, "일/월", "정제기 비교", "제공", USER, "", "int")
inp("mh", "월 비교 기준 시간 (정제기 1대)", f"=B{r-1}*24", "h/월", "정제기 전용 — 검사·충진 인력에 자동 적용 안 함", "수식", "수식", "", "h")
inp("yh", "연간 달력시간", "=365*24", "h/년", "연간 연속 환산", "수식", "수식", "정지·보수 미반영 상한", "int")

sec("B. 공정시간 (충진 전까지 57 h 구성)")
inp("t_iqc", "수입검사 IQC", 2, "h", "57 h에 포함 · 원료 Lot 단위 (Batch 공통 시 중복 배정 금지)", "제공", USER57, "NMR·ICP·IC", "h")
inp("t_prep", "준비·투입", 2, "h/Batch", "57 h에 포함 · 정제기 점유", "제공", USER57, "", "h")
inp("t_ref", "정제", 45, "h/Batch", "57 h에 포함 · 이번 자료 정제시간 기준", "제공", USER57, "57 − 2 − 2 − 6 − 2 = 45", "h")
inp("t_pqctr", "공정검사 PQC + 제품 이송 (합계)", 6, "h/Batch", "두 공정 합계 — 각각 6 h 적용 금지", "제공", USER57, "그림에서는 구분", "h")
inp("t_fqc", "제품검사 FQC", 2, "h/Batch", "57 h에 포함", "제공", USER57, "", "h")
inp("t57", "충진 전까지 합계 (제공)", 57, "h/Batch", "2+2+45+6+2", "제공", USER, "", "h")
inp("t57c", "충진 전까지 합계 (검증 수식)", f"={R('t_iqc')}+{R('t_prep')}+{R('t_ref')}+{R('t_pqctr')}+{R('t_fqc')}", "h/Batch", "제공값 57과 일치해야 함", "수식", "수식", "", "h")
inp("t_ref49", "참고: 생산팀 대화 중 정제 49 h 언급", 49, "h", "이번 자료 미적용", "구두", ORAL, "45 h 기준 적용 (사용자 결정)", "h")
inp("t_trocc", "제품 이송 중 정제기 점유시간", None, "h/Batch", "PQC·이송 6 h 중 정제기 점유분", "미확인", "-", "미입력 시 점유 = 준비 2 + 정제 45 (하한)")
inp("t_fqcw", "검사 대기시간 (FQC 등)", None, "h/Batch", "작업시간과 구분", "미확인", "-")
inp("t_clean", "세척·전환", None, "h/회", "", "미확인", "-")
inp("t_maint", "보수·정지 (정제기)", None, "h/월", "월별 입력은 04 시트 F열", "미확인", "-")
inp("t_mix", "Mix·리사이클 준비 (설비시간 사용)", None, "h/Batch", "To-be 2대 운전 시 보조 작업", "미확인", "-", "52.4톤과 47.2톤 차이 요인 후보 — 확정 손실 아님")

sec("C. 충진 · OQC (필터는 충진에 포함)")
inp("t_5g", "5 Gal 충진 (글로브 박스)", 2, "h/병", "하이닉스·CXMT 20 kg/병", "제공", USER, "근거 없이 1 h로 변경 금지", "h")
inp("n_5g_b", "한 Batch 5 Gal 충진 병수 (참고)", 9, "병/Batch", "약 9병", "제공", USER, "9 × 20 = 180 kg", "int")
inp("t_5g_b", "한 Batch 5 Gal 충진시간", f"={R('t_5g')}*{R('n_5g_b')}", "h/Batch", "병당 h × 병수 (약 9병 18 h)", "수식", "수식", "", "h")
inp("chk_5g", "5 Gal 정합 확인 (병수×병당)", f"={R('n_5g_b')}*{R('t_5g')}", "h/Batch", "18 h와 일치", "수식", "수식", "", "h")
inp("t_oqc5", "5 Gal OQC·출하 작업", 2, "h/9병", "9병 합계 2 h", "제공", USER, "물류 운송시간 별도", "h")
inp("t_oqc2", "200 L OQC·출하 작업", 2, "h/용기", "1용기 2 h", "제공", USER, "물류 운송시간 별도", "h")
inp("t_ez_m", "이지켐 200 L 수동 충진", 8, "h/용기", "현재 수동", "제공", USER, "", "h")
inp("t_hs_m", "한솔 200 L 수동 충진", None, "h/용기", "'27.1~6", "미확인", "-", "이지켐 8 h를 확정값으로 복제하지 않음")
inp("t_ez_a", "이지켐 ARS 충진", None, "h/용기", "ARS 전환 후", "미확인", "-")
inp("t_hs_a", "한솔 ARS 충진", None, "h/용기", "'27.7~", "미확인", "-")
inp("t_common", "공통 시간 적용 가정 (200 L 비교용)", 8, "h/용기", "미확인 고객에 이지켐 수동 8 h 공통 적용 시 비교", "가정", "비교용 가정", "확정값 아님 — 별도 열에만 사용", "h")
inp("wh_day", "근무일 환산 기준 (일 근무시간)", 8, "h/일", "18 h ÷ 8 h = 2.25 근무일", "제공", USER, "현장: 5 Gal 충진 약 2일 소요 설명", "h")
inp("tgt_pack", "충진·포장 목표 (개선 필요사항)", "1~1.5일", "일", "미달성 — 개선 과제", "계획", USER, "병당 2 h 변경·달성 실적으로 표시하지 않음")

sec("D. 고객 · 용기 · 충진 방식")
inp("kg_5g", "5 Gal 충진량 (하이닉스·CXMT)", 20, "kg/병", "5 Gal 약 19 L · 글로브 박스 유지", "제공", USER, "", "kg0")
inp("kg_ez", "이지켐 200 L 충진량", 140, "kg/용기", "현재 수동 → 향후 ARS", "제공", USER, "", "kg0")
inp("kg_hs", "한솔 200 L 충진량", 150, "kg/용기", "'27.1~6 수동 / '27.7~ ARS", "제공", USER, "", "kg0")
inp("ars_inst", "ARS 설치 (기존 계획)", dt.date(2026, 12, 31), "날짜", "200 L 전용 · 5 Gal 미적용", "계획", "기존 자료 ('26 설치 → '27 운영)", "설치일 미확정 — 2026년 내")
inp("ars_hs", "한솔 ARS 적용 시작", dt.date(2027, 7, 1), "날짜", "한솔", "제공", USER, "ARS 운영 시점 2027년 7월 · 2028년 전월 ARS")
inp("ars_ez", "이지켐 ARS 전환 월", dt.date(2027, 7, 1), "날짜", "이지켐", "계획", "사용자 지시 (ARS 운영 시점 2027년 7월)",
    "이지켐·한솔 동일 '27.7~ ARS · '27.1~6 수동")

sec("E. 생산량 · 수율 · 잔량")
inp("kg_b", "Batch당 생산량 (환산 기준)", 190, "kg/Batch", "기본 환산 기준", "제공", USER, "구두 수치로 임의 변경 금지", "kg0")
inp("kg_9b", "9병 충진량", f"={R('n_5g_b')}*{R('kg_5g')}", "kg", "9 × 20", "수식", "수식", "", "kg0")
inp("kg_res", "190 kg − 9병 차이 (잔량 후보)", f"={R('kg_b')}-B{r-1}", "kg/Batch", "잔량·추가 충진·재고 처리 확인", "수식", "수식", "자동 손실 처리 금지", "kg0")
inp("o_crude", "구두: Crude 투입량", 200, "kg/Batch", "약", "구두", ORAL, "", "kg0")
inp("o_rec_lo", "구두: 순수 회수율 하한", 0.75, "%", "약 75~80%", "구두", ORAL, "", "pct")
inp("o_rec_hi", "구두: 순수 회수율 상한", 0.80, "%", "", "구두", ORAL, "", "pct")
inp("o_cases", "구두: 조건별 Batch량 언급", "160~170 / 약 220 / 평균 195~200 (197)", "kg", "리사이클·추가 투입으로 보충", "구두", ORAL, "하나의 확정값으로 통합 안 함")
inp("o_cases2", "구두: 135 kg · 120 kg 사례", "135 / 120", "kg", "집계 범위 다름", "구두", ORAL, "")
inp("o_ez", "구두: 이지켐향 추가 투입 제한", "색도 관리로 추가 투입 제한", "-", "이지켐", "구두", ORAL, "'순도 기준 없음'·'전체 적합'으로 확정 안 함")
inp("y_in_new", "Batch 신규 Crude 투입량 (실측)", None, "kg/Batch", "수율 계산", "미확인", "-")
inp("y_in_rec", "Batch 재순환 투입량 (실측)", None, "kg/Batch", "", "미확인", "-")
inp("y_good", "Batch 양품량 (실측)", None, "kg/Batch", "", "미확인", "-", "불량·재작업 손실 중복 차감 금지")
inp("kg_b_rf", "개선 후 Batch당 생산량", None, "kg/Batch", "'27.7~", "미확인", "-", "미입력 시 190 kg 유지로 계산·표시")

sec("F. Capa. (연간 환산 기준)")
inp("capa_now", "현재 연간 Capa.", 26.2, "t/년", "기준값 유지", "제공", USER, "시간·가동률·수율 조정해 맞추지 않음", "t")
inp("capa_add", "개선 증가분", 21, "t/년", "'27.7~'28.6 최초 12개월", "제공", USER, "정제기 2대 활용", "t")
inp("capa_rf", "개선 후 연간 환산 Capa.", f"={R('capa_now')}+{R('capa_add')}", "t/년", "≠ 2027 실제 생산량", "수식", "수식", "ARS 효과 미가산", "t")
inp("capa_x2", "참고: 현재 Capa. 단순 2배", f"={R('capa_now')}*2", "t/년", "적용 안 함 — Mix·리사이클 보조 작업 등", "수식", "수식", "52.4 − 47.2 차이를 확정 손실로 표시 안 함", "t")
inp("capa_tadd", "105℃ 예상 증가분", 3, "t/년", "2028 검토안", "검토안", USER, "미확정·검증·승인 필요", "t")
inp("capa_t", "105℃ 적용 시 산술 Capa.", f"={R('capa_rf')}+{R('capa_tadd')}", "t/년", "검토안", "수식", "수식", "", "t")
inp("target", "목표 Capa.", 50, "t/년", "2028년 이후 50톤 수준", "계획", USER, "", "t")
inp("max11", "과거 월 최대 생산 Batch", 11, "Batch/월", "히스토리 최대", "제공", USER, "", "int")
inp("temp_new", "검토 정제온도", 105, "℃", "2028 검토안", "검토안", USER, "열 안정성·Dimer·Unknown impurity·Yield", "int")
inp("t105", "105℃ 적용 시작 월", None, "날짜", "04 시트 월별 반영", "미확인", "-", "미정 — 확정 시 입력")
inp("int_105", "105℃ 적용 시 설비별 Batch 간격", None, "h/Batch", "", "미확인", "-")

sec("G. 일정 · 정제기 2대 시차 운전")
inp("con_s", "공사 시작", dt.date(2027, 2, 1), "날짜", "2027년 2~5월", "계획", USER)
inp("con_e", "공사 종료", dt.date(2027, 5, 31), "날짜", "", "계획", USER)
inp("trial_s", "시운전", dt.date(2027, 6, 1), "날짜", "2027년 6월", "계획", USER)
inp("rf_s", "개선 적용 생산 시작", dt.date(2027, 7, 1), "날짜", "", "계획", USER)
inp("f12_e", "최초 12개월 종료", dt.date(2028, 6, 30), "날짜", "2027.7~2028.6", "계획", USER)
inp("h_con", "공사 중 정제기 가용시간", None, "h/월", "2027.2~5", "미확인", "-", "0 또는 정상값 임의 입력 안 함")
inp("b_trial", "시운전 양품·출하 승인 Batch", None, "Batch", "2027.6", "미확인", "-")
inp("stab_m", "초기 안정화 기간", None, "개월", "'27.7~", "미확인", "-")
inp("int_rf", "개선 후 설비별 Batch 투입 간격", None, "h/Batch", "정제기 1·2 각각", "미확인", "-", "45 h → 22.5 h 단축으로 계산하지 않음")
inp("offset", "정제기 2 기동 시차 (정제기 1 투입 시작 기준)", 30, "h", "Gantt 예시", "가정", "운영 개념 설명용 가정", "미확정 — 동시 기동·24 h 고정 아님")

sec("H. 후공정 자원 (정제기 720 h와 별도)")
inp("gb_n", "글로브 박스 수", None, "대", "5 Gal", "미확인", "-")
inp("gb_h", "글로브 박스 충진 가용 근무시간", None, "h/월", "교대·인력", "미확인", "-")
inp("m200_h", "200 L 수동/ARS 가용시간", None, "h/월", "", "미확인", "-")
inp("qc_h", "검사(QC) 가용시간", None, "h/월", "", "미확인", "-")
inp("tank_n", "Product Tank 수", None, "기", "", "미확인", "-")
inp("tank_cap", "Product Tank 용량", None, "kg", "", "미확인", "-")
inp("sc_d", "참고 시나리오: 월 근무일", 22, "일/월", "03 시트 비교용", "가정", "참고 가정", "확정값 아님")
inp("sc1", "참고 ①: 1교대 일 근무", 8, "h/일", "", "가정", "참고 가정", "", "h")
inp("sc2", "참고 ②: 2교대 일 근무", 16, "h/일", "", "가정", "참고 가정", "", "h")
inp("sc3", "참고 ③: 연속 근무", 24, "h/일", "30일", "가정", "참고 가정", "", "h")

sec("I. 과거 일정 분석 설정")
inp("an_s", "분석 시작 Batch", 39, "#", "6/2", "가정", "분석 설정", "5월 일부 제외", "int")
inp("an_e", "분석 종료 Batch", 92, "#", "10/29", "가정", "분석 설정", "", "int")
inp("roll_k", "롤링 검토 간격 수", 4, "간격", "", "가정", "분석 설정", "", "int")
inp("basis", "현재 조건 Batch 간격 기준 (1~4)", 3, "선택", "1=이론 45 h, 2=과거 최대 11, 3=관측 추정, 4=26.2톤 역산", "가정", "분석 설정", "04·05 현재 조건 생산 가능량", "int")
inp("hist_y", "과거 Batch 일정 연도", 2026, "년", "원자료 연도 미표기", "가정", "첨부 PPT가 2026 출하와 비교", "", "int")

sec("J. 재고 · 품질")
inp("inv0", "기초재고 (2026-01-01)", None, "kg", "36개월 연속 이월", "미확인", "-", "미입력 시 재고 계산 보류 (0 초기화 안 함)")
inp("tgt_hx", "하이닉스 목표재고", None, "kg", "우선 확보", "미확인", "-", "재고 수준 — 매년 반복 가산 안 함")
inp("tgt_all", "전체 목표 기말재고", None, "kg", "", "미확인", "-")
inp("y_all", "정제 Batch 양품률", None, "%", "", "미확인", "-", "미입력 시 100% 상한값으로 표시")
inp("y_hx", "하이닉스 합격률", None, "%", "", "미확인", "-")
inp("y_cx", "CXMT 합격률", None, "%", "", "미확인", "-")
inp("y_ez", "이지켐 합격률 (색도 포함)", None, "%", "", "미확인", "-", "색도 관리 중요 — 전체 적합 확정 안 함")
inp("y_hs", "한솔 합격률", None, "%", "", "미확인", "-")
inp("prio", "하이닉스 외 고객 우선순위", None, "-", "CXMT·이지켐·한솔", "미확인", "-", "미입력 시 잔여량을 출하 비율로 배분 (우선순위 가정 없음)")

last_input_row = r - 1
wsI.auto_filter.ref = f"A5:G{last_input_row}"

# ---- shipment grid (36 months)
r += 1
wsI.cell(row=r, column=1, value="고객별 월별 출하계획 (kg) — 2026 제시분 · 2027 계획 · 2028 가정").font = font("EB002C", True, 11)
r += 1
header(wsI, r, ["고객", "연도"] + [f"{m}월" for m in range(1, 13)] + ["연간", "실적/계획 구분", "출처", "비고"])
r += 1
SHIP = {}
hx26 = [1600, 1420, 1440, 1580, 1200, 1200, 1360, 1420, 1380, 1480, 1480, 1420]
ez26 = [0] * 8 + [560] * 4


def grow(label, year, vals, kind, cls, src, note, key, nf="kg"):
    global r
    style_cell(wsI.cell(row=r, column=1, value=label), "text")
    style_cell(wsI.cell(row=r, column=2, value=year), "text", "int")
    for m in range(12):
        v = vals[m] if vals is not None else None
        style_cell(wsI.cell(row=r, column=3 + m, value=v), kind if not (isinstance(v, str) and v.startswith("=")) else "calc", nf)
    style_cell(wsI.cell(row=r, column=15, value=f'=IF(COUNT(C{r}:N{r})=0,"미제시",SUM(C{r}:N{r}))'), "key", nf)
    for col, v in ((16, cls), (17, src), (18, note)):
        style_cell(wsI.cell(row=r, column=col, value=v), "text")
    SHIP[key] = r
    r += 1


for yr in (2026, 2027, 2028):
    sheet_ = {2026: "P3_과거Batch_2026", 2027: "P4_2027_월별", 2028: "P5_2028_월별"}[yr]
    cls_ = {2026: "원자료 실적/계획 구분 미확인", 2027: "계획", 2028: "가정 (확정 수요 아님)"}[yr]
    for cust, lab in (("hx", "하이닉스"), ("cx", "CXMT"), ("ez", "이지켐"), ("hs", "한솔")):
        refs = [ship_cell(yr, cust, m) for m in range(1, 13)]
        nt_ = "미제시 시 공란 — 합계 제외" if (yr == 2026 and cust in ("cx", "hs")) else ""
        grow(lab, yr, [f'=IF(ISBLANK({x}),"",{x})' for x in refs], "link", cls_, f"{sheet_} 입력 (연결)", nt_, (yr, cust))
for yr in (2026, 2027, 2028):
    rows_ = [SHIP[(yr, k)] for k in ("hx", "cx", "ez", "hs")]
    style_cell(wsI.cell(row=r, column=1, value=f"{yr} 합계" + (" (제시분: 하이닉스+이지켐)" if yr == 2026 else "")), "text", bold=True)
    style_cell(wsI.cell(row=r, column=2, value=yr), "text", "int")
    for m in range(13):
        col = CL(3 + m)
        fx = "=" + "+".join(f"N({col}{x})" for x in rows_) if m < 12 else f"=SUM(C{r}:N{r})"
        style_cell(wsI.cell(row=r, column=3 + m, value=fx), "key", "kg")
    for col, v in ((16, "합계"), (17, "수식"), (18, {2026: "CXMT·한솔 미제시분 미포함", 2027: "검증 43,400 kg", 2028: "검증 45,760 kg"}[yr])):
        style_cell(wsI.cell(row=r, column=col, value=v), "text")
    SHIP[(yr, "tot")] = r
    r += 1
for yr, tgt in ((2026, 19220), (2027, 43400), (2028, 45760)):
    style_cell(wsI.cell(row=r, column=1, value=f"검증: {yr} 합계 = {tgt:,} kg"), "text")
    style_cell(wsI.cell(row=r, column=15, value=f'=IF(ABS(O{SHIP[(yr, "tot")]}-{tgt})<0.001,"일치","불일치")'), "key")
    wsI.conditional_formatting.add(f"O{r}", CellIsRule(operator="equal", formula=['"불일치"'], fill=red, font=redf))
    r += 1
widths(wsI, {"A": 38, "B": 13, "C": 11, "D": 32, "E": 9, "F": 26, "G": 36, "H": 10, "I": 10, "J": 10, "K": 10,
             "L": 10, "M": 10, "N": 10, "O": 12, "P": 18, "Q": 22, "R": 30})
wsI.freeze_panes = "B6"


def ship_ref(year, cust, month):
    return f"'01_Inputs'!${CL(2 + month)}${SHIP[(year, cust)]}"


# ======================================================== 02_Batch_Raw
wsB = wb.create_sheet("02_Batch_Raw")
title(wsB, "02_Batch_Raw | 과거 Batch 생산계획 원자료 · 대조 · 간격",
      "A~N열 원자료(수정 금지) · O열 이후 분석 수식. 표시일자는 날짜만 있는 생산계획 — 실제 운전 시각·실적 아님.")
legend(wsB, 3)
raw_hdr = ["No", "원본 파일", "출처 위치", "Batch 번호", "표시일자", "실제 시각\n(미확인)", "표시일자 의미\n(확인란)", "계획/실적",
           "원자료 색상\n(분포도)", "색상 의미\n(확인란)", "노트 표기\n(원본·수정본)", "분포도 판독\n일(日)", "사용자 지시\n목록 일자", "대조 결과"]
ana_hdr = ["월", "간격 분석\n포함", "번호 연속", "중복", "이전과 간격\n(일)", "월 경계", "명목 간격\n(h)", "최소 경과\n(h)",
           "최대 경과\n(h)", "롤링 k간격\n누적(일)", "롤링 평균\n최대 경과 (h)", "정제기 점유 정합"]
header(wsB, 5, raw_hdr, 1, "hdr", 42); header(wsB, 5, ana_hdr, 15, "hdr2", 42)
L = [(37, 5, 23), (38, 5, 26), (39, 6, 2), (40, 6, 5), (41, 6, 8), (42, 6, 11), (43, 6, 14), (44, 6, 17), (45, 6, 20), (46, 6, 23),
     (47, 6, 25), (48, 6, 27), (49, 6, 28), (50, 7, 1), (51, 7, 3), (52, 7, 6), (53, 7, 8), (54, 7, 11), (55, 7, 14), (56, 7, 17),
     (57, 7, 20), (58, 7, 23), (59, 7, 26), (60, 7, 29), (61, 8, 1), (62, 8, 4), (63, 8, 7), (64, 8, 10), (65, 8, 13), (66, 8, 16),
     (67, 8, 19), (68, 8, 22), (69, 8, 25), (70, 8, 28), (71, 8, 31), (72, 9, 3), (73, 9, 5), (74, 9, 7), (75, 9, 9), (76, 9, 11),
     (77, 9, 14), (78, 9, 19), (79, 9, 23), (80, 9, 26), (81, 9, 28), (82, 10, 1), (83, 10, 4), (84, 10, 7), (85, 10, 10),
     (86, 10, 12), (87, 10, 15), (88, 10, 18), (89, 10, 20), (90, 10, 23), (91, 10, 26), (92, 10, 29)]
n1 = Presentation(SRC1).slides[2]; n2 = Presentation(SRC2).slides[2]
notes1 = n1.notes_slide.notes_text_frame.text; notes2 = n2.notes_slide.notes_text_frame.text
chart = {}
for sh in n1.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().isdigit():
        n = int(sh.text_frame.text.strip()); y = (sh.top + sh.height / 2) / 914400
        if 37 <= n <= 92 and 1.5 < y < 2.75 and abs(sh.width / 914400 - 0.154) < 0.01:
            x = (sh.left + sh.width / 2) / 914400
            chart[n] = (round(1 + (x - 1.012) / 0.17110), {"FF7900": "주황", "7F7F7F": "회색"}.get(str(sh.fill.fore_color.rgb), "?"))
B0 = 6
NMAX = 130
BL = B0 + (NMAX - 37)
Ld = {n: (m, d) for n, m, d in L}
G_, G1_ = P3_GRID, P3_GRID_FIRST
for i, n in enumerate(range(37, NMAX + 1)):
    rr = B0 + i
    m, d = Ld.get(n, (None, None))
    if m:
        tok = f"{m}/{d} #{n}"
        nt = f"{m}/{d}" if (tok in notes1 and tok in notes2) else ("원본만" if tok in notes1 else ("수정본만" if tok in notes2 else "미발견"))
        meta = [chart.get(n, (None, "?"))[1], None, nt, chart.get(n, (None,))[0], dt.date(2026, m, d)]
    else:
        meta = [None, None, "-", None, None]
    efx = (f'=IF(COUNTIF({G_},D{rr})=0,"",IF(COUNTIF({G_},D{rr})>1,"중복 입력",SUMPRODUCT(({G_}=D{rr})*'
           f'DATE({R("hist_y")},ROW({G_})-ROW({G1_})+5,COLUMN({G_})-COLUMN({G1_})+1))))')
    vals = [i + 1, SRC_NAME if m else "P3 달력 입력", "P3 달력 ← 원본 3장 노트·분포도" if m else "P3_과거Batch_2026 달력", n, efx, None, None,
            "생산계획 (원자료) — 실적 미검증"] + meta
    for j, v in enumerate(vals):
        style_cell(wsB.cell(row=rr, column=1 + j, value=v), "link" if j == 4 else ("in" if j in (3, 8, 10, 11, 12) else ("unk" if j in (5, 6, 9) else "text")))
    wsB.cell(row=rr, column=5).number_format = NF["d"]; wsB.cell(row=rr, column=13).number_format = NF["d"]
    nfx = (f'=IF(AND(ISNUMBER(E{rr}),E{rr}=M{rr},DAY(E{rr})=L{rr},K{rr}=MONTH(E{rr})&"/"&DAY(E{rr})),"일치","차이 확인")' if m
           else f'=IF(ISNUMBER(E{rr}),"신규 입력","")')
    style_cell(wsB.cell(row=rr, column=14, value=nfx))
    occ = f"({R('t_prep')}+{R('t_ref')}+N({R('t_trocc')}))"
    f = {15: f'=IF(ISNUMBER(E{rr}),MONTH(E{rr}),"")', 16: f"=IF(AND(ISNUMBER(S{rr}),D{rr}>{R('an_s')},D{rr}<={R('an_e')}),1,0)",
         17: '="-"' if i == 0 else f'=IF(AND(ISNUMBER(E{rr}),NOT(ISNUMBER(E{rr-1}))),"확인","")',
         18: f'=IF(E{rr}="중복 입력","중복","-")',
         19: None if i == 0 else f'=IF(AND(ISNUMBER(E{rr}),ISNUMBER(E{rr-1})),E{rr}-E{rr-1},"")',
         20: None if i == 0 else f'=IF(ISNUMBER(S{rr}),IF(MONTH(E{rr})<>MONTH(E{rr-1}),"월 경계",""),"")',
         21: None if i == 0 else f'=IF(ISNUMBER(S{rr}),S{rr}*24,"")', 22: None if i == 0 else f'=IF(ISNUMBER(S{rr}),MAX(0,(S{rr}-1)*24),"")',
         23: None if i == 0 else f'=IF(ISNUMBER(S{rr}),(S{rr}+1)*24,"")',
         24: (f'=IF(AND(ISNUMBER(E{rr}),ROW()-{B0}>={R("roll_k")}),IF(ISNUMBER(INDEX($E${B0}:$E${BL},ROW()-{B0}+1-{R("roll_k")})),'
              f'E{rr}-INDEX($E${B0}:$E${BL},ROW()-{B0}+1-{R("roll_k")}),""),"")'),
         25: f'=IF(ISNUMBER(X{rr}),(X{rr}+1)*24/{R("roll_k")},"")',
         26: f'=IF(ISNUMBER(Y{rr}),IF(Y{rr}<{occ},"확인 필요: 점유 하한 미만",""),"")'}
    for col, fx in f.items():
        style_cell(wsB.cell(row=rr, column=col, value=fx), "calc", {19: "int", 21: "int", 22: "int", 23: "int", 24: "int", 25: "h"}.get(col))
wsB.auto_filter.ref = f"A5:Z{BL}"; wsB.freeze_panes = "E6"
widths(wsB, {"A": 5, "B": 26, "C": 26, "D": 8, "E": 12, "F": 10, "G": 12, "H": 18, "I": 10, "J": 10, "K": 11, "L": 10, "M": 12,
             "N": 10, "O": 6, "P": 8, "Q": 8, "R": 7, "S": 10, "T": 9, "U": 9, "V": 9, "W": 9, "X": 10, "Y": 12, "Z": 20})
S0 = BL + 3
wsB.cell(row=S0 - 1, column=1, value="월별 집계 · 간격 분포 (간격은 뒤 Batch의 월에 귀속, 분석 범위 내)").font = font("EB002C", True, 11)
header(wsB, S0, ["월", "번호 범위", "번호 개수", "원자료 Total", "차이\n(번호−Total)", "번호×190\n(kg)", "Total×190\n(kg)", "첫 표시일",
                 "마지막 표시일", "간격 수", "평균 간격\n(일)", "중앙값\n(일)", "최소", "최대", "1일", "2일", "3일", "4일", "5일", "6일 이상",
                 "월 경계\n간격 수", "비고"], 1, "hdr", 40)
E_ = f"$E${B0}:$E${BL}"; O_ = f"$O${B0}:$O${BL}"; P_ = f"$P${B0}:$P${BL}"; S_ = f"$S${B0}:$S${BL}"; D_ = f"$D${B0}:$D${BL}"; T_ = f"$T${B0}:$T${BL}"
SUMROW = {}
MONTHS_ = [(5, "5월 (일부)", "#37·#38만 제공 — 주 추정 제외")] + [(m_, f"{m_}월", "") for m_ in range(6, 13)]
for k, (m, lab, note) in enumerate(MONTHS_):
    rr = S0 + 1 + k; SUMROW[m] = rr
    tot = f'=IF(ISNUMBER({P3_TOTAL[m]}),{P3_TOTAL[m]},"")'
    cells = {1: lab, 2: f'="#"&_xlfn.MINIFS({D_},{O_},{m})&"~#"&_xlfn.MAXIFS({D_},{O_},{m})', 3: f"=COUNTIF({O_},{m})", 4: tot,
             5: f'=IF(ISNUMBER(D{rr}),C{rr}-D{rr},"원자료 Total 없음")', 6: f"=C{rr}*{R('kg_b')}", 7: f'=IF(ISNUMBER(D{rr}),D{rr}*{R("kg_b")},"-")',
             8: f"=_xlfn.MINIFS({E_},{O_},{m})", 9: f"=_xlfn.MAXIFS({E_},{O_},{m})", 10: f"=COUNTIFS({O_},{m},{P_},1)",
             11: f'=IF(J{rr}>0,AVERAGEIFS({S_},{O_},{m},{P_},1),"-")', 13: f'=IF(J{rr}>0,_xlfn.MINIFS({S_},{O_},{m},{P_},1),"-")',
             14: f'=IF(J{rr}>0,_xlfn.MAXIFS({S_},{O_},{m},{P_},1),"-")', 20: f'=COUNTIFS({O_},{m},{P_},1,{S_},">=6")',
             21: f'=COUNTIFS({O_},{m},{P_},1,{T_},"월 경계")', 22: note}
    for dd, col in zip(range(1, 6), range(15, 20)):
        cells[col] = f"=COUNTIFS({O_},{m},{P_},1,{S_},{dd})"
    for col, v in cells.items():
        kind = "link" if col == 4 else ("text" if col in (1, 22) else "calc")
        style_cell(wsB.cell(row=rr, column=col, value=v), kind, {6: "kg0", 7: "kg0", 8: "md", 9: "md", 11: "b"}.get(col))
    wsB[f"L{rr}"] = ArrayFormula(f"L{rr}", f'=IF(J{rr}>0,MEDIAN(IF(({O_}={m})*({P_}=1),{S_})),"-")'); style_cell(wsB[f"L{rr}"], "calc", "b")
rt = S0 + 1 + len(MONTHS_); SUMROW["tot"] = rt; r6, r10 = SUMROW[6], SUMROW[12]
tc = {1: "6~12월 합계", 2: '="#"&' + f'_xlfn.MINIFS({D_},{O_},">=6")&"~#"&_xlfn.MAXIFS({D_},{O_},"<=12")', 3: f"=SUM(C{r6}:C{r10})",
      4: f"=SUM(D{r6}:D{r10})", 5: f"=C{rt}-D{rt}", 6: f"=SUM(F{r6}:F{r10})", 7: f"=SUM(G{r6}:G{r10})", 8: f"=H{r6}", 9: f"=I{r10}",
      10: f"=SUM(J{r6}:J{r10})", 11: f"=AVERAGEIFS({S_},{P_},1)", 13: f"=_xlfn.MINIFS({S_},{P_},1)", 14: f"=_xlfn.MAXIFS({S_},{P_},1)",
      21: f"=SUM(U{r6}:U{r10})", 22: "5→6월 경계(7일)는 분석 범위 밖"}
for col in range(15, 21):
    tc[col] = f"=SUM({CL(col)}{r6}:{CL(col)}{r10})"
for col, v in tc.items():
    style_cell(wsB.cell(row=rt, column=col, value=v), "text" if col in (1, 22) else "key", {6: "kg0", 7: "kg0", 8: "md", 9: "md", 11: "b"}.get(col), True)
wsB[f"L{rt}"] = ArrayFormula(f"L{rt}", f"=MEDIAN(IF({P_}=1,{S_}))"); style_cell(wsB[f"L{rt}"], "key", "b")
wsB.conditional_formatting.add(f"E{r6}:E{rt}", FormulaRule(formula=[f"AND(ISNUMBER(E{r6}),E{r6}<>0)"], fill=red, font=redf))
wsB.conditional_formatting.add(f"C{r6}:D{r10}", FormulaRule(formula=[f"AND(ISNUMBER($D{r6}),$C{r6}<>$D{r6})"], fill=red, font=redf))
wsB.conditional_formatting.add(f"N{B0}:N{BL}", FormulaRule(formula=[f'N{B0}="차이 확인"'], fill=red, font=redf))
wsB.conditional_formatting.add(f"Z{B0}:Z{BL}", FormulaRule(formula=[f'LEFT(Z{B0},2)="확인"'], fill=red, font=redf))
wsB.conditional_formatting.add(f"S{B0}:S{BL}", FormulaRule(formula=[f"AND(ISNUMBER(S{B0}),S{B0}<=2)"], fill=PatternFill("solid", fgColor="FFEB9C")))
rc = rt + 2
CHK = {}
for k, (lab, fx) in enumerate([("대조 '차이 확인' 건수", f'=COUNTIF($N${B0}:$N${BL},"차이 확인")'), ("번호 불연속", f'=COUNTIF($Q${B0}:$Q${BL},"확인")'),
                               ("중복 번호", f'=COUNTIF($R${B0}:$R${BL},"중복")'), ("Total≠번호 월 수", f"=SUMPRODUCT(--ISNUMBER(E{r6}:E{r10}),--(E{r6}:E{r10}<>0))"),
                               ("정제기 점유 하한 미만 롤링 구간", f'=COUNTIF($Z${B0}:$Z${BL},"확인*")')]):
    style_cell(wsB.cell(row=rc + k, column=1, value=lab), "text"); wsB.merge_cells(start_row=rc + k, start_column=1, end_row=rc + k, end_column=4)
    style_cell(wsB.cell(row=rc + k, column=5, value=fx), "key", "int"); CHK[k] = f"'02_Batch_Raw'!$E${rc+k}"
for k, t in enumerate(["※ 원본 이미지 5개는 첨부되지 않음 → 원본·수정본 발표자 노트 + 원본 분포도 좌표 + 사용자 지시 목록 대조.",
                       "※ D일 간격 실제 경과 = (D−1)×24 ~ (D+1)×24 h. 2일 = 24~72 h — 48 h로 단정하지 않음.",
                       "※ 정제기 점유 정합: 롤링 k간격 평균 최대 경과가 점유 하한(준비 2 + 정제 45 + 이송 점유 입력)보다 짧으면 표시.",
                       "※ 번호 개수 ≠ 양품 완료 Batch. 원자료 Total 집계 대상·색상 의미·표시일자(투입/완료) 의미 확인 필요."]):
    wsB.cell(row=rc + 6 + k, column=1, value=t).font = font(GRAY, italic=True)

# ======================================================== 03_Capacity_Model
wsC = wb.create_sheet("03_Capacity_Model")
title(wsC, "03_Capacity_Model | 공정시간 · 정제기 점유 · 이론/과거/관측 · 2대 시차 운전 · 후공정 자원",
      "Lead Time(입고→출하)과 정제기 투입 간격을 구분. 충진은 다음 Batch 정제와 병렬 — 충진시간을 정제시간에 더하지 않음.")
legend(wsC, 3)
widths(wsC, {"A": 50, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 46})
C = {}
row = 5


def cblock(label):
    global row
    row += 1
    wsC.cell(row=row, column=1, value=label).font = font("EB002C", True, size=11)
    row += 1


def cline(key, label, fx, unit, basis, kind="calc", nf="h2"):
    global row
    style_cell(wsC.cell(row=row, column=1, value=label), "text")
    style_cell(wsC.cell(row=row, column=2, value=fx), kind, nf)
    style_cell(wsC.cell(row=row, column=3, value=unit), "text")
    wsC.merge_cells(start_row=row, start_column=4, end_row=row, end_column=8)
    style_cell(wsC.cell(row=row, column=4, value=basis), "note")
    if key:
        C[key] = f"'03_Capacity_Model'!$B${row}"
    row += 1


cblock("1. 경로별 한 Batch 확인 공정시간 (단순 합계 — 달력 납기 아님)")
header(wsC, row, ["공정", "5 Gal 약 9병\n(h)", "이지켐 200 L\n1용기 수동 (h)", "한솔 200 L\n1용기 (h)", "", "", "", "비고"]); row += 1
lt0 = row
steps = [("수입검사 IQC", "t_iqc", "t_iqc", "t_iqc", "원료 Lot 단위"), ("준비·투입", "t_prep", "t_prep", "t_prep", "정제기 점유"),
         ("정제", "t_ref", "t_ref", "t_ref", "정제기 점유"), ("PQC + 제품 이송 (합계)", "t_pqctr", "t_pqctr", "t_pqctr", "합계 6 h"),
         ("제품검사 FQC", "t_fqc", "t_fqc", "t_fqc", "대기 별도"), ("충진 (필터 포함)", "t_5g_b", "t_ez_m", "t_hs_m", "한솔 수동 시간 미확인"),
         ("OQC·출하 작업", "t_oqc5", "t_oqc2", "t_oqc2", "운송 별도")]
for lab, a, b_, c_, note in steps:
    style_cell(wsC.cell(row=row, column=1, value=lab), "text")
    for col, k in ((2, a), (3, b_), (4, c_)):
        style_cell(wsC.cell(row=row, column=col, value=f'=IF(ISNUMBER({R(k)}),{R(k)},"미확인")'), "link", "h")
    style_cell(wsC.cell(row=row, column=8, value=note), "note")
    row += 1
lt1 = row - 1
style_cell(wsC.cell(row=row, column=1, value="충진 전까지 소계 (IQC~FQC)"), "text", bold=True)
for col in (2, 3, 4):
    style_cell(wsC.cell(row=row, column=col, value=f"=SUM({CL(col)}{lt0}:{CL(col)}{lt0+4})"), "key", "h")
C["lt57"] = f"'03_Capacity_Model'!$B${row}"; row += 1
style_cell(wsC.cell(row=row, column=1, value="확인시간 단순 합계"), "text", bold=True)
for col in (2, 3, 4):
    fx = (f'=SUM({CL(col)}{lt0}:{CL(col)}{lt1})&IF(COUNTIF({CL(col)}{lt0}:{CL(col)}{lt1},"미확인")>0," + 미확인","")' if col == 4
          else f"=SUM({CL(col)}{lt0}:{CL(col)}{lt1})")
    style_cell(wsC.cell(row=row, column=col, value=fx), "key", "h")
C["lt_5g"] = f"'03_Capacity_Model'!$B${row}"; C["lt_ez"] = f"'03_Capacity_Model'!$C${row}"; C["lt_hs"] = f"'03_Capacity_Model'!$D${row}"
style_cell(wsC.cell(row=row, column=8, value="77 h · 67 h — 실제 달력 납기 아님 (대기·근무시간 미반영)"), "note"); row += 1
cline("el_5g", "5 Gal 충진 18 h의 근무일 환산", f"={R('t_5g_b')}/{R('wh_day')}", "근무일", "18 ÷ 8 = 2.25 근무일 (현장 설명: 약 2일) — 작업시간 ≠ 경과시간", nf="h2")
cline("el_5g_h", "1교대 기준 충진 경과시간 (달력)", f"=ROUNDUP({C['el_5g']},0)*24", "h", "근무일 올림 × 24 h — 참고", nf="h")

cblock("2. 정제기 점유시간 vs 다음 Batch 투입 간격")
cline("occ", "정제기 점유 확인분 (준비·투입 + 정제)", f"={R('t_prep')}+{R('t_ref')}", "h/Batch", "IQC·FQC는 정제기 점유 아님", "key", "h")
cline("occ_hi", "점유 상한 참고 (+ PQC·이송 6 h 전부)", f"={C['occ']}+{R('t_pqctr')}", "h/Batch", "이송 중 점유분 미확인 (01_Inputs B)", nf="h")
cline("occ_in", "적용 점유 (확인분 + 이송 점유 입력)", f"={C['occ']}+N({R('t_trocc')})", "h/Batch", "미입력 시 하한", nf="h")

cblock("3. ① 정제시간만 적용한 이론 비교값 (실제 생산능력 아님)")
cline("th_t", "정제시간", f"={R('t_ref')}", "h/Batch", "45 h", "link", "h")
cline("th_bm", "720 ÷ 45", f"={R('mh')}/{C['th_t']}", "Batch/월", "준비·전환·이송·세척·대기·보수 제외")
cline("th_bi", "월내 완료 정수 Batch", f"=INT({C['th_bm']})", "Batch/월", "", "key", "int")
cline("th_carry", "월말 진행 중 경과 (이월)", f"={R('mh')}-{C['th_bi']}*{C['th_t']}", "h", "", nf="h")
cline("th_kgm", "월 생산량 (정수 × 190)", f"={C['th_bi']}*{R('kg_b')}", "kg/월", "16 × 190 = 3,040", "key", "kg0")
cline("th_occ_bm", "참고: 점유 확인분 47 h 기준 720 ÷ 47", f"={R('mh')}/{C['occ']}", "Batch/월", "준비·투입 포함 시", nf="b")
cline("th_by", "참고: 연간 달력 연속 INT(8,760 ÷ 45)", f"=INT({R('yh')}/{C['th_t']})", "Batch/년", "정지·보수 0 가정 상한", nf="int")
cline("th_kgy", "참고: 연간 달력 연속 생산량", f"={C['th_by']}*{R('kg_b')}", "kg/년", "", nf="kg0")

cblock("4. ② 과거 월 최대 11 Batch")
cline("mx_b", "과거 월 최대", f"={R('max11')}", "Batch/월", "", "link", "int")
cline("mx_ref", "순수 정제시간 합계 11 × 45", f"={C['mx_b']}*{C['th_t']}", "h/월", "", nf="h")
cline("mx_gap", "720 − 495", f"={R('mh')}-{C['mx_ref']}", "h/월", "점유·전환·대기·보수·월 경계 확인 — 전부 손실/충진으로 분류 안 함", "key", "h")
cline("mx_occ", "참고: 점유 확인분 합계 11 × 47", f"={C['mx_b']}*{C['occ']}", "h/월", "", nf="h")
cline("mx_int", "720 ÷ 11 배분 간격", f"={R('mh')}/{C['mx_b']}", "h/Batch", "30일 월 기준", nf="h2")
cline("mx_kgm", "월 생산량 11 × 190", f"={C['mx_b']}*{R('kg_b')}", "kg/월", "", "key", "kg0")
cline("mx_kgy12", "매월 11회 반복 참고 (× 12)", f"={C['mx_kgm']}*12", "kg/년", "25.08톤 — 참고값", "key", "kg0")
cline("mx_kgy", "참고: 달력 연속 INT(8,760 ÷ 65.45) × 190", f"=INT({R('yh')}/{C['mx_int']})*{R('kg_b')}", "kg/년", "", nf="kg0")

cblock("5. ③ 과거 일정 관측 추정 (02_Batch_Raw)")
Bq = "'02_Batch_Raw'!"
E_r = f"{Bq}$E${B0}:$E${BL}"; D_r = f"{Bq}$D${B0}:$D${BL}"; P_r = f"{Bq}$P${B0}:$P${BL}"
cline("ob_d0", "주 범위 시작 표시일", f"=INDEX({E_r},MATCH({R('an_s')},{D_r},0))", "날짜", "6/2 #39", "link", "yyyy-mm-dd")
cline("ob_d1", "주 범위 종료 표시일", f"=INDEX({E_r},MATCH({R('an_e')},{D_r},0))", "날짜", "10/29 #92", "link", "yyyy-mm-dd")
cline("ob_n_dates", "표시일자 수", f"={R('an_e')}-{R('an_s')}+1", "개", "54", nf="int")
cline("ob_days", "경과 일수", f"={C['ob_d1']}-{C['ob_d0']}", "일", "", nf="int")
cline("ob_n", "간격 수", f"=SUM({P_r})", "간격", "53 (02 시트 포함 표시 합)", nf="int")
cline("ob_int", "일자차 환산 평균 간격", f"={C['ob_days']}*24/{C['ob_n']}", "h/Batch", "경과일×24 ÷ 간격 수", "key", "h2")
cline("ob_lo", "범위 하한 (시각 미상)", f"=({C['ob_days']}-1)*24/{C['ob_n']}", "h/Batch", "", nf="h2")
cline("ob_hi", "범위 상한", f"=({C['ob_days']}+1)*24/{C['ob_n']}", "h/Batch", "", nf="h2")
cline("ob_med", "개별 간격 중앙값", f"='02_Batch_Raw'!$L${rt}*24", "h", "72 h", "link", "h")
cline("ob_share3", "3일 간격 비중", f"='02_Batch_Raw'!$Q${rt}/'02_Batch_Raw'!$J${rt}", "%", "", "link", "pct")
cline("ob_vs45", "평균 간격 − 정제 45 h", f"={C['ob_int']}-{C['th_t']}", "h/Batch", "준비·이송·세척·대기·IQC 대기·계획 여유 등 — 구성 미확인", nf="h2")
cline("ob_vs47", "평균 간격 − 점유 확인분 47 h (비점유 시간)", f"={C['ob_int']}-{C['occ']}", "h/Batch", "", "key", "h2")
cline("ob_b30", "월 Batch (30일)", f"=30*24/{C['ob_int']}", "Batch/월", "", nf="b")
cline("ob_b31", "월 Batch (31일)", f"=31*24/{C['ob_int']}", "Batch/월", "11 Batch 월(7·8·10월 31일)과 정합", nf="b")
cline("ob_kgm", "월 생산량 (30일, 소수 Batch)", f"={C['ob_b30']}*{R('kg_b')}", "kg/월", "", nf="kg0")
cline("ob_by", "연간 Batch (달력 연속 정수)", f"=INT({R('yh')}/{C['ob_int']})", "Batch/년", "", nf="int")
cline("ob_kgy", "연간 생산량 (달력 연속)", f"={C['ob_by']}*{R('kg_b')}", "kg/년", "현재 1대 · 정지·보수 추가 미반영", "key", "kg0")
cline("ob_alt37", "참고: #37 포함 평균", f"=({C['ob_d1']}-INDEX({E_r},MATCH(37,{D_r},0)))*24/({R('an_e')}-37)", "h/Batch", "5/26→6/2 7일 포함", nf="h2")
cline("ob_dense", "점유 하한 미만 롤링 구간 수", f"={CHK[4]}", "구간", "45 h 기준 — 57 h 가정 시 확인 필요였던 구간도 재평가", "link", "int")

cblock("6. 현재 Capa. 26.2톤 환산 (기준값 유지)")
cline("cp_kg", "현재 Capa.", f"={R('capa_now')}*1000", "kg/년", "", "link", "kg0")
cline("cp_b", "Batch 상당", f"={C['cp_kg']}/{R('kg_b')}", "Batch/년", "26,200 ÷ 190 = 137.895", nf="0.000")
cline("cp_bround", "정수 Batch (반올림)", f"=ROUND({C['cp_b']},0)", "Batch/년", "", nf="int")
cline("cp_kground", "정수 Batch 생산량", f"={C['cp_bround']}*{R('kg_b')}", "kg/년", "138 × 190 = 26,220 → 26.22 → 26.2톤", nf="kg0")
cline("cp_bm", "월 평균 Batch 상당", f"={C['cp_b']}/12", "Batch/월", "", nf="b")
cline("cp_int", "역산 평균 간격 (8,760 h)", f"={R('yh')}/{C['cp_b']}", "h/Batch", "참고", "key", "h2")

cblock("7. 생산능력 비교 (정제기 1대 · 190 kg/Batch)")
header(wsC, row, ["구분", "Batch 간격\n(h)", "월 Batch\n(30일)", "월 생산\n(kg)", "연 환산\n(t, 참고)", "26.2톤 대비\n(t)", "", "산출 조건"]); row += 1
CMP0 = row
for lab, fx, yr_fx, basis in (("① 이론 (정제 45 h만)", f"={C['th_t']}", f"={C['th_kgm']}*12/1000", "720 ÷ 45 = 16 · 준비·전환·이송·세척·대기·보수 제외"),
                              ("② 과거 월 최대 11 Batch", f"={C['mx_int']}", f"={C['mx_kgy12']}/1000", "정제 495 h + 225 h · 매월 11회 반복 참고"),
                              ("③ 과거 일정 관측 추정", f"={C['ob_int']}", f"={C['ob_kgy']}/1000", "6/2~10/29 · 149일 ÷ 53간격 · 달력 연속"),
                              ("현재 Capa. 26.2톤 환산", f"={C['cp_int']}", f"={R('capa_now')}", "26,200 ÷ 190 = 137.9 Batch 상당")):
    style_cell(wsC.cell(row=row, column=1, value=lab), "text", bold=True)
    for j, v in enumerate([fx, f"=30*24/B{row}", f"=C{row}*{R('kg_b')}", yr_fx, f"=E{row}-{R('capa_now')}"]):
        style_cell(wsC.cell(row=row, column=2 + j, value=v), "key" if j == 3 else "calc", ["h2", "b", "kg0", "t", "t"][j])
    style_cell(wsC.cell(row=row, column=8, value=basis), "note"); row += 1
cline("basis_int", "04·05 적용 현재 조건 간격 (01_Inputs 선택)", f"=CHOOSE({R('basis')},{C['th_t']},{C['mx_int']},{C['ob_int']},{C['cp_int']})", "h/Batch", "기본 3 = 관측 추정", "key", "h2")

cblock("8. To-be 정제기 2대 시간차 운전 (정제 45 h 단축 아님)")
cline("x2", "참고: 현재 26.2톤 단순 2배", f"={R('capa_x2')}", "t/년", "적용 안 함", "link", "t")
cline("x2_gap", "단순 2배 − 개선 후 47.2톤", f"={C['x2']}-{R('capa_rf')}", "t/년", "Mix·리사이클 준비 등 보조 작업·공유 자원 가능성 — 확정 손실 아님", "key", "t")
cline("x2_ratio", "47.2 ÷ 52.4", f"={R('capa_rf')}/{C['x2']}", "%", "", nf="pct")
cline("rf_b", "47.2톤 Batch 상당 (190 kg)", f"={R('capa_rf')}*1000/{R('kg_b')}", "Batch/년", "", nf="b1")
cline("rf_int_all", "2대 합산 평균 투입 간격 (역산)", f"={R('yh')}/{C['rf_b']}", "h/Batch", "공장 전체 — 설비별 정제시간 아님", nf="h2")
cline("rf_int_each", "설비별 등가 투입 간격 (역산)", f"={R('n_new')}*{R('yh')}/{C['rf_b']}", "h/Batch", "현재 1대 26.2톤 등가 63.5 h와 비교", "key", "h2")
cline("rf_int_diff", "설비별 등가 간격 − 현재 26.2톤 등가", f"={C['rf_int_each']}-{C['cp_int']}", "h/Batch", "보조 작업 시간 가능성 (구성 미확인)", nf="h2")
cline("off", "정제기 2 기동 시차 (예시 가정)", f"={R('offset')}", "h", "운영 개념 설명용 가정 · 미확정", "link", "h")
cline("g_int", "Gantt 설비별 간격", f'=IF(ISNUMBER({R("int_rf")}),{R("int_rf")},ROUND({C["rf_int_each"]},1))', "h/Batch", "입력 없으면 47.2톤 역산 참고값", "est", "h")
header(wsC, row, ["Gantt (h, 정제기 1 IQC 시작 = 0)", "시작", "종료", "자원", "", "", "", "비고"]); row += 1
G0 = row
gl = []
for b in range(2):
    for unit, off in (("정제기 1", "0"), ("정제기 2", C["off"])):
        base = f"({off}+{b}*{C['g_int']})"
        gl.append((f"{unit} B{b+1} 준비·투입", f"={base}+{R('t_iqc')}", f"=B{{r}}+{R('t_prep')}", unit, ""))
        gl.append((f"{unit} B{b+1} 정제", "=C{p}", f"=B{{r}}+{R('t_ref')}", unit, "45 h"))
        gl.append((f"{unit} B{b+1} PQC·이송·FQC", "=C{p}", f"=B{{r}}+{R('t_pqctr')}+{R('t_fqc')}", "검사·Tank", "6 + 2 h"))
        gl.append((f"{unit} B{b+1} 5 Gal 충진", "=C{p}", f"=B{{r}}+{R('t_5g_b')}", "글로브 박스 (공용)", "연속 작업 18 h"))
for k, (lab, s, e, res, note) in enumerate(gl):
    rr = G0 + k
    style_cell(wsC.cell(row=rr, column=1, value=lab), "text")
    style_cell(wsC.cell(row=rr, column=2, value=s.replace("{p}", str(rr - 1)).replace("{r}", str(rr))), "calc", "h")
    style_cell(wsC.cell(row=rr, column=3, value=e.replace("{p}", str(rr - 1)).replace("{r}", str(rr))), "calc", "h")
    style_cell(wsC.cell(row=rr, column=4, value=res), "text"); style_cell(wsC.cell(row=rr, column=8, value=note), "note")
GN = G0 + len(gl) - 1
C["g0"], C["gn"] = G0, GN
for j in range(41):
    c = wsC.cell(row=G0 - 1, column=10 + j, value=j * 6); c.font = font(GRAY, size=7); c.alignment = CENTER
    wsC.column_dimensions[CL(10 + j)].width = 3.2
    for rr in range(G0, GN + 1):
        cc = wsC.cell(row=rr, column=10 + j, value=f'=IF(AND({CL(10+j)}${G0-1}+6>$B{rr},{CL(10+j)}${G0-1}<$C{rr}),1,"")')
        cc.font = font("FFFFFF", size=6)
rng = f"J{G0}:{CL(50)}{GN}"
for kw, col in (("정제", "FF7900"), ("준비", "F4B183"), ("PQC", "2E75B6"), ("충진", "548235")):
    wsC.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(J{G0}=1,ISNUMBER(SEARCH("{kw}",$A{G0})))'], fill=PatternFill("solid", fgColor=col), font=Font(color=col)))
row = GN + 1
cline("fill_gap_min", "충진 투입 최소 간격 (두 설비 교차)", f"=MIN({C['off']},{C['g_int']}-{C['off']})", "h", "충진이 한꺼번에 몰리지 않도록 시차 배치", nf="h")

cblock("9. 후공정 작업량 vs 가용시간 (정제기 720 h 자동 적용 안 함)")
MQ = "'04_Monthly_2026_2028'!"
cline("gb27", "5 Gal 충진 작업 2027 연간", f"=SUM({MQ}AB18:AB29)", "h/년", "(하이닉스+CXMT) ÷ 20 × 2 — OQC 별도", "link", "h")
cline("gb28", "5 Gal 충진 작업 2028 연간", f"=SUM({MQ}AB30:AB41)", "h/년", "", "link", "h")
cline("gbmax", "5 Gal 월 최대 (2027~2028)", f"=MAX({MQ}AB18:AB41)", "h/월", "", "link", "h")
cline("c200_27", "200 L 용기 2027 연간 (이지켐+한솔)", f"=SUM({MQ}AD18:AE29)", "용기", "", "link", "int")
cline("c200_28", "200 L 용기 2028 연간", f"=SUM({MQ}AD30:AE41)", "용기", "", "link", "int")
header(wsC, row, ["가용시간 시나리오 (글로브 박스)", "일 근무 (h)", "근무일", "월 가용 (h)", "5 Gal 최대\n부하율", "5 Gal 여유\n(h)", "", "비고"]); row += 1
PP0 = row
for lab, hk, dk, note in (("입력값", None, None, "01_Inputs H — 미확인"), ("참고 ① 1교대", "sc1", "sc_d", "가정"),
                          ("참고 ② 2교대", "sc2", "sc_d", "가정"), ("참고 ③ 연속 (30일)", "sc3", None, "가정")):
    style_cell(wsC.cell(row=row, column=1, value=lab), "text", bold=True)
    vals = ["-", "-", f'=IF(ISNUMBER({R("gb_h")}),{R("gb_h")},"미확인")'] if hk is None else [f"={R(hk)}", f"={R(dk)}" if dk else "=30", f"=B{row}*C{row}"]
    vals += [f'=IF(ISNUMBER(D{row}),{C["gbmax"]}/D{row},"미확인")', f'=IF(ISNUMBER(D{row}),D{row}-{C["gbmax"]},"미확인")']
    for j, v in enumerate(vals):
        style_cell(wsC.cell(row=row, column=2 + j, value=v), "calc", ["h", "int", "h", "pct", "h"][j])
    style_cell(wsC.cell(row=row, column=8, value=note), "note"); row += 1
C["pp0"] = PP0

cblock("10. Product Tank 회전 (Tank 1기 가정 시: FQC·충진 완료 ≤ 다음 제품 도착)")
header(wsC, row, ["5 Gal Batch 충진 18 h 경과", "일 근무\n(h)", "충진 경과\n(h)", "Tank 점유\n(FQC+충진)", "여유 vs As-is\n관측 간격", "여유 vs To-be\n2대 합산 간격", "", "판단"]); row += 1
TK0 = row
for lab, hk in (("참고 ① 1교대", "sc1"), ("참고 ② 2교대", "sc2"), ("참고 ③ 연속", "sc3")):
    style_cell(wsC.cell(row=row, column=1, value=lab), "text", bold=True)
    vals = [f"={R(hk)}", f"={R('t_5g_b')}/(B{row}/24)", f"={R('t_fqc')}+N({R('t_fqcw')})+C{row}", f"={C['ob_int']}-D{row}",
            f"={C['rf_int_all']}-D{row}", f'=IF(F{row}<0,"To-be Tank·충진 병목 가능","여유")']
    for j, v in enumerate(vals):
        style_cell(wsC.cell(row=row, column=2 + j, value=v), "calc", ["h", "h", "h", "h", "h", None][j])
    row += 1
C["tk0"] = TK0

cblock("11. 생산량 · 수율 · 잔량 (구두 수치는 확정값으로 통합하지 않음)")
cline("res9", "190 kg − 9병(180 kg)", f"={R('kg_res')}", "kg/Batch", "잔량·추가 충진·재고 처리 확인 — 자동 손실 처리 안 함", "link", "kg0")
cline("mix_ez", "혼합 경로 예: 이지켐 1용기 후 잔량", f"={R('kg_b')}-{R('kg_ez')}", "kg", "190 − 140 = 50 kg → 5 Gal 2병 + 10 kg 또는 보관", nf="kg0")
cline("mix_hs", "혼합 경로 예: 한솔 1용기 후 잔량", f"={R('kg_b')}-{R('kg_hs')}", "kg", "190 − 150 = 40 kg → 5 Gal 2병", nf="kg0")
cline("o_lo", "구두 참고: Crude 200 kg × 75%", f"={R('o_crude')}*{R('o_rec_lo')}", "kg", "순수 회수 — 리사이클·추가 투입 전", "oral", "kg0")
cline("o_hi", "구두 참고: Crude 200 kg × 80%", f"={R('o_crude')}*{R('o_rec_hi')}", "kg", "", "oral", "kg0")
cline("y_new", "실측 수율 (양품 ÷ 신규 Crude)", f'=IF(AND(ISNUMBER({R("y_good")}),ISNUMBER({R("y_in_new")})),{R("y_good")}/{R("y_in_new")},"보류")', "%", "실측 입력 시 계산", nf="pct")
cline("y_tot", "실측 수율 (양품 ÷ 총 투입 = 신규 + 재순환)", f'=IF(AND(ISNUMBER({R("y_good")}),ISNUMBER({R("y_in_new")}),ISNUMBER({R("y_in_rec")})),{R("y_good")}/({R("y_in_new")}+{R("y_in_rec")}),"보류")', "%", "", nf="pct")
wsC.freeze_panes = "B5"

# ======================================================== 04_Monthly_2026_2028
wsM = wb.create_sheet("04_Monthly_2026_2028")
title(wsM, "04_Monthly_2026_2028 | 36개월 생산 가능량 · 4개 고객 출하 · 충진 · 재고",
      "생산 가능 Batch = INT((설비 가용시간 합계 + 이월) ÷ 설비별 간격). 미확인 조건은 '미확인/보류' — 0으로 치환하지 않음.")
legend(wsM, 3)
cols = [("A", "연월"), ("B", "연도"), ("C", "월 일수"), ("D", "운전 조건"), ("E", "정제기 수"), ("F", "정지·보수\n입력 (h)"),
        ("G", "정제기 가용\n합계 (h)"), ("H", "설비별 Batch\n간격 (h)"), ("I", "이월 (h)"), ("J", "생산 가능\n완료 Batch"),
        ("K", "월말 진행\n이월 (h)"), ("L", "과거계획\n번호 개수"), ("M", "과거계획\nTotal"), ("N", "확정 생산 Batch\n(실적 입력)"),
        ("O", "적용 Batch"), ("P", "적용 기준"), ("Q", "Batch당\n생산량 (kg)"), ("R", "정제 생산량\n(kg)"), ("S", "출하 가능 양품\n(kg)"),
        ("T", "참고 Capa.\n월 환산 (kg)"), ("U", "하이닉스\n(kg)"), ("V", "CXMT\n(kg)"), ("W", "이지켐\n(kg)"), ("X", "한솔\n(kg)"),
        ("Y", "출하 합계\n(kg)"), ("Z", "필요 Batch\n(÷190)"), ("AA", "5 Gal 병 상당\n(병)"), ("AB", "5 Gal 충진\n작업 (h)"),
        ("AC", "5 Gal OQC\n(h)"), ("AD", "이지켐\n용기"), ("AE", "한솔\n용기"), ("AF", "이지켐\n방식"), ("AG", "한솔\n방식"),
        ("AH", "이지켐 충진\n(h)"), ("AI", "한솔 충진\n(h)"), ("AJ", "200 L 충진\n확인분 (h)"), ("AK", "200 L 공통 8 h\n적용 가정 (h)"),
        ("AL", "200 L OQC\n(h)"), ("AM", "글로브 박스\n가용 (h)"), ("AN", "글로브 박스\n여유 (h)"), ("AO", "참고 Capa.\n− 출하 (kg)"),
        ("AP", "양품 생산\n− 출하 (kg)"), ("AQ", "누적 증감\n(기초재고 제외)"), ("AR", "기초재고"), ("AS", "공급 가능량"),
        ("AT", "① 하이닉스\n배정"), ("AU", "② 하이닉스\n목표재고"), ("AV", "③ CXMT\n배정"), ("AW", "③ 이지켐\n배정"), ("AX", "③ 한솔\n배정"),
        ("AY", "출하 부족"), ("AZ", "기말재고"), ("BA", "목표\n기말재고"), ("BB", "필요\n생산량"), ("BC", "여유·부족\n(기말−목표)"),
        ("BD", "계산 상태"), ("BE", "미확인 항목")]
header(wsM, 5, [c[1] for c in cols], 1, "hdr", 54)
M0 = 6; MN = M0 + 35
for k in range(36):
    rr = M0 + k; y = 2026 + k // 12; m = k % 12 + 1; first = k == 0; p = rr - 1
    f = {}
    f["A"] = dt.date(y, m, 1); f["B"] = f"=YEAR(A{rr})"; f["C"] = f"=DAY(EOMONTH(A{rr},0))"
    f["D"] = (f'=IF(A{rr}<{R("con_s")},"현재 (1대)",IF(A{rr}<={R("con_e")},"공사",IF(A{rr}<{R("trial_s")},"현재 (1대)",IF(A{rr}<{R("rf_s")},"시운전",'
              f'IF(AND(ISNUMBER({R("t105")}),A{rr}>={R("t105")}),"개선 (2대·105℃)","개선 (2대)")))))')
    f["E"] = f'=IF(LEFT(D{rr},2)="개선",{R("n_new")},{R("n_now")})'
    f["F"] = None
    f["G"] = f'=IF(D{rr}="공사",IF(ISNUMBER({R("h_con")}),{R("h_con")},"미확인"),IF(D{rr}="시운전","시운전",E{rr}*C{rr}*24-N(F{rr})))'
    f["H"] = (f'=IF(OR(D{rr}="현재 (1대)",D{rr}="공사"),{C["basis_int"]},IF(D{rr}="개선 (2대)",IF(ISNUMBER({R("int_rf")}),{R("int_rf")},"미확인"),'
              f'IF(D{rr}="개선 (2대·105℃)",IF(ISNUMBER({R("int_105")}),{R("int_105")},"미확인"),"-")))')
    f["I"] = "=0" if first else f"=IF(AND(ISNUMBER(K{p}),D{p}=D{rr}),K{p},0)"
    f["J"] = (f'=IF(D{rr}="시운전",IF(ISNUMBER({R("b_trial")}),{R("b_trial")},"미확인"),'
              f'IF(AND(ISNUMBER(G{rr}),ISNUMBER(H{rr})),INT((G{rr}+I{rr})/H{rr}),"미확인"))')
    f["K"] = f'=IF(AND(ISNUMBER(G{rr}),ISNUMBER(H{rr}),ISNUMBER(J{rr})),G{rr}+I{rr}-J{rr}*H{rr},"")'
    f["L"] = (f'=IF(COUNTIFS({E_r},">="&A{rr},{E_r},"<="&EOMONTH(A{rr},0))=0,"",COUNTIFS({E_r},">="&A{rr},{E_r},"<="&EOMONTH(A{rr},0)))')
    f["M"] = (f"=IF(ISNUMBER('02_Batch_Raw'!$D${SUMROW[m]}),'02_Batch_Raw'!$D${SUMROW[m]},\"\")" if (y == 2026 and m in SUMROW) else '=""')
    f["N"] = None
    f["O"] = f"=IF(ISNUMBER(N{rr}),N{rr},J{rr})"
    f["P"] = f'=IF(ISNUMBER(N{rr}),"실적 입력",IF(ISNUMBER(J{rr}),"추정 생산 가능량","미확인"))'
    f["Q"] = f'=IF(LEFT(D{rr},2)="개선",IF(ISNUMBER({R("kg_b_rf")}),{R("kg_b_rf")},{R("kg_b")}),{R("kg_b")})'
    f["R"] = f'=IF(ISNUMBER(O{rr}),O{rr}*Q{rr},"미확인")'
    f["S"] = f'=IF(ISNUMBER(R{rr}),R{rr}*IF(ISNUMBER({R("y_all")}),{R("y_all")},1),"미확인")'
    f["T"] = f'=IF(A{rr}<{R("rf_s")},{R("capa_now")},IF(D{rr}="개선 (2대·105℃)",{R("capa_t")},{R("capa_rf")}))*1000/12'
    for col, cu in (("U", "hx"), ("V", "cx"), ("W", "ez"), ("X", "hs")):
        sr = ship_ref(y, cu, m)
        f[col] = f'=IF(ISNUMBER({sr}),{sr},"미제시")'
    f["Y"] = f"=SUM(U{rr}:X{rr})"
    f["Z"] = f"=Y{rr}/Q{rr}"
    f["AA"] = f"=(N(U{rr})+N(V{rr}))/{R('kg_5g')}"
    f["AB"] = f"=AA{rr}*{R('t_5g')}"
    f["AC"] = f"=AA{rr}/{R('n_5g_b')}*{R('t_oqc5')}"
    f["AD"] = f"=N(W{rr})/{R('kg_ez')}"
    f["AE"] = f"=N(X{rr})/{R('kg_hs')}"
    f["AF"] = f'=IF(AD{rr}=0,"-",IF(A{rr}>={R("ars_ez")},"ARS","수동"))'
    f["AG"] = f'=IF(AE{rr}=0,"-",IF(A{rr}>={R("ars_hs")},"ARS","수동"))'
    f["AH"] = f'=IF(AD{rr}=0,0,IF(AF{rr}="수동",AD{rr}*{R("t_ez_m")},IF(ISNUMBER({R("t_ez_a")}),AD{rr}*{R("t_ez_a")},"미확인")))'
    f["AI"] = (f'=IF(AE{rr}=0,0,IF(AG{rr}="수동",IF(ISNUMBER({R("t_hs_m")}),AE{rr}*{R("t_hs_m")},"미확인"),'
               f'IF(ISNUMBER({R("t_hs_a")}),AE{rr}*{R("t_hs_a")},"미확인")))')
    f["AJ"] = f'=IF(AND(ISNUMBER(AH{rr}),ISNUMBER(AI{rr})),AH{rr}+AI{rr},"부분 "&TEXT(N(AH{rr})+N(AI{rr}),"0")&" h + 미확인")'
    f["AK"] = f"=(AD{rr}+AE{rr})*{R('t_common')}"
    f["AL"] = f"=(AD{rr}+AE{rr})*{R('t_oqc2')}"
    f["AM"] = f'=IF(ISNUMBER({R("gb_h")}),{R("gb_h")},"미확인")'
    f["AN"] = f'=IF(ISNUMBER(AM{rr}),AM{rr}-AB{rr},"미확인")'
    f["AO"] = f"=T{rr}-Y{rr}"
    f["AP"] = f'=IF(ISNUMBER(S{rr}),S{rr}-Y{rr},"미확인")'
    f["AQ"] = f'=IF(ISNUMBER(AP{rr}),AP{rr},"보류")' if first else f'=IF(AND(ISNUMBER(AP{rr}),ISNUMBER(AQ{p})),AQ{p}+AP{rr},"보류")'
    f["AR"] = f'=IF(ISNUMBER({R("inv0")}),{R("inv0")},"미입력")' if first else f'=IF(ISNUMBER(AZ{p}),AZ{p},"보류")'
    f["AS"] = f'=IF(AND(ISNUMBER(AR{rr}),ISNUMBER(S{rr})),AR{rr}+S{rr},"보류")'
    f["AT"] = f'=IF(ISNUMBER(AS{rr}),MIN(AS{rr},U{rr}),"보류")'
    f["AU"] = f'=IF(ISNUMBER(AS{rr}),MIN(AS{rr}-AT{rr},N({R("tgt_hx")})),"보류")'
    oth = f"(N(V{rr})+N(W{rr})+N(X{rr}))"
    for col, src in (("AV", "V"), ("AW", "W"), ("AX", "X")):
        f[col] = f'=IF(ISNUMBER(AS{rr}),IF({oth}=0,0,MIN(N({src}{rr}),(AS{rr}-AT{rr}-AU{rr})*N({src}{rr})/{oth})),"보류")'
    f["AY"] = f'=IF(ISNUMBER(AS{rr}),Y{rr}-AT{rr}-AV{rr}-AW{rr}-AX{rr},"보류")'
    f["AZ"] = f'=IF(ISNUMBER(AS{rr}),AS{rr}-AT{rr}-AV{rr}-AW{rr}-AX{rr},"보류")'
    f["BA"] = f'=IF(ISNUMBER({R("tgt_all")}),{R("tgt_all")},"미입력")'
    f["BB"] = f'=IF(AND(ISNUMBER(AR{rr}),ISNUMBER(BA{rr})),Y{rr}+BA{rr}-AR{rr},"보류")'
    f["BC"] = f'=IF(AND(ISNUMBER(AZ{rr}),ISNUMBER(BA{rr})),AZ{rr}-BA{rr},"보류")'
    f["BD"] = f'=IF(ISNUMBER(AZ{rr}),"계산 가능",IF(ISNUMBER(AP{rr}),"부분 (재고 입력 필요)","보류 (생산 조건 미확인)"))'
    f["BE"] = "=" + "&".join([
        f'IF(D{rr}="공사",IF(ISNUMBER({R("h_con")}),"","공사 중 가용시간·"),"")',
        f'IF(D{rr}="시운전",IF(ISNUMBER({R("b_trial")}),"","시운전 양품·"),"")',
        f'IF(D{rr}="개선 (2대)",IF(ISNUMBER({R("int_rf")}),"","설비별 간격·"),"")',
        f'IF(D{rr}="개선 (2대·105℃)",IF(ISNUMBER({R("int_105")}),"","105℃ 간격·"),"")',
        f'IF(ISNUMBER(N{rr}),"","실적·")', f'IF(ISNUMBER({R("y_all")}),"","양품률(100% 상한)·")',
        f'IF(COUNTIF(U{rr}:X{rr},"미제시")>0,"고객 물량 미제시·","")',
        f'IF(ISNUMBER(AJ{rr}),"","200 L 충진시간·")', f'IF(ISNUMBER({R("gb_h")}),"","글로브 박스 가용·")',
        f'IF(ISNUMBER({R("inv0")}),"","기초재고·")'])
    nfs = {"A": "ym", "B": "int", "C": "int", "E": "int", "F": "h", "G": "h", "H": "h2", "I": "h", "J": "int", "K": "h", "L": "int", "M": "int",
           "N": "int", "O": "int", "Q": "kg0", "R": "kg0", "S": "kg0", "T": "kg", "U": "kg", "V": "kg", "W": "kg", "X": "kg", "Y": "kg",
           "Z": "b1", "AA": "b1", "AB": "h", "AC": "h", "AD": "b1", "AE": "b1", "AH": "h", "AI": "h", "AJ": "h", "AK": "h", "AL": "h",
           "AM": "h", "AN": "h", "AO": "kg", "AP": "kg", "AQ": "kg", "AR": "kg", "AS": "kg", "AT": "kg", "AU": "kg", "AV": "kg", "AW": "kg",
           "AX": "kg", "AY": "kg", "AZ": "kg", "BA": "kg", "BB": "kg", "BC": "kg"}
    for col, _ in cols:
        c = wsM[f"{col}{rr}"]; c.value = f.get(col)
        kind = "unk" if col in ("F", "N") else ("in" if col == "A" else ("link" if col in ("U", "V", "W", "X", "L", "M") else "calc"))
        if col in ("J", "S", "Y", "AZ"):
            kind = "key"
        style_cell(c, kind, nfs.get(col))
SR = {2026: MN + 1, 2027: MN + 2, 2028: MN + 3}
for yr, rr in SR.items():
    a = M0 + (yr - 2026) * 12; b = a + 11
    style_cell(wsM.cell(row=rr, column=1, value=f"{yr} 합계"), "text", bold=True)
    for col in ("J", "O", "R", "S"):
        wsM[f"{col}{rr}"] = (f'=IF(COUNT({col}{a}:{col}{b})=12,SUM({col}{a}:{col}{b}),"부분합계 "&TEXT(SUM({col}{a}:{col}{b}),"#,##0")'
                             f'&" ("&COUNT({col}{a}:{col}{b})&"/12개월)")'); style_cell(wsM[f"{col}{rr}"], "key", "kg0")
    for col in ("T", "U", "W", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AK", "AL", "AO") + (("L", "M") if yr == 2026 else ()):
        wsM[f"{col}{rr}"] = f"=SUM({col}{a}:{col}{b})"; style_cell(wsM[f"{col}{rr}"], "key", nfs.get(col))
    for col in ("V", "X"):
        wsM[f"{col}{rr}"] = f'=IF(COUNT({col}{a}:{col}{b})=0,"미제시",SUM({col}{a}:{col}{b}))'; style_cell(wsM[f"{col}{rr}"], "key", "kg")
    for col in ("AH", "AI"):
        wsM[f"{col}{rr}"] = f'=IF(COUNTIF({col}{a}:{col}{b},"미확인")>0,"부분 "&TEXT(SUM({col}{a}:{col}{b}),"0")&" h + 미확인",SUM({col}{a}:{col}{b}))'
        style_cell(wsM[f"{col}{rr}"], "key", "h")
MX = {2026: MN + 4, 2027: MN + 5, 2028: MN + 6}
for yr, rr in MX.items():
    a = M0 + (yr - 2026) * 12; b = a + 11
    style_cell(wsM.cell(row=rr, column=1, value=f"{yr} 월 최대"), "text", bold=True)
    for col in ("Y", "AA", "AB", "AC", "AD", "AE", "AK", "AL"):
        wsM[f"{col}{rr}"] = f"=MAX({col}{a}:{col}{b})"; style_cell(wsM[f"{col}{rr}"], "calc", nfs.get(col))
wsM.freeze_panes = "E6"; wsM.auto_filter.ref = f"A5:BE{MN}"
for i in range(1, len(cols) + 1):
    wsM.column_dimensions[CL(i)].width = 10.5
for k, v in {"A": 9, "D": 15, "P": 14, "AJ": 16, "BD": 18, "BE": 60}.items():
    wsM.column_dimensions[k].width = v
for col in ("AN", "AO", "AP", "AQ", "BC"):
    wsM.conditional_formatting.add(f"{col}{M0}:{col}{MN}", FormulaRule(formula=[f"AND(ISNUMBER({col}{M0}),{col}{M0}<0)"], fill=red, font=redf))
wsM.conditional_formatting.add(f"AY{M0}:AY{MN}", FormulaRule(formula=[f"AND(ISNUMBER(AY{M0}),AY{M0}>0.001)"], fill=red, font=redf))
wsM.conditional_formatting.add(f"G{M0}:BD{MN}", FormulaRule(formula=[f'OR(G{M0}="미확인",G{M0}="보류",G{M0}="미입력",G{M0}="미제시",LEFT(G{M0},2)="부분")'],
                                                           fill=PatternFill("solid", fgColor="EDEDED"), font=Font(color="C55A11", italic=True)))
for k, t in enumerate(["※ 생산 가능 Batch는 정제기 능력 추정 — 실제 계획·실적 아님. 실적은 N열 입력 시 우선 적용. 2027.12 기말재고 → 2028.1 기초재고로 자동 연결.",
                       "※ 배정: ① 하이닉스 출하 → ② 하이닉스 목표재고 → ③ CXMT·이지켐·한솔은 잔여량을 출하 비율로 배분 (우선순위 미확인 — 01_Inputs J).",
                       "※ 기말재고 = 기초재고 + 출하 가능 양품 생산량 − 출하량(배정분) · 필요 생산량 = 출하량 + 목표 기말재고 − 기초재고.",
                       "※ 5 Gal 병 상당 = (하이닉스+CXMT) ÷ 20 — 실제 출하는 20 kg 정수 용기로 배분하고 연간 물량 유지. 200 L: 이지켐 140 kg / 한솔 150 kg.",
                       "※ 200 L 충진 확인분은 미확인 고객 시간 제외. '공통 8 h 적용 가정' 열은 비교용 가정 — 한솔 확정값 아님.",
                       "※ 참고 Capa. 월 환산 = 연간 환산 Capa. ÷ 12 (2027.7~ 47.2톤) — 실제 달력연도 생산량 아님."]):
    wsM.cell(row=MX[2028] + 2 + k, column=1, value=t).font = font(GRAY, italic=True)

# ======================================================== 05_Reflux_Scenarios
wsR = wb.create_sheet("05_Reflux_Scenarios")
title(wsR, "05_Reflux_Scenarios | 공사 · 시운전 · 개선(정제기 2대) · 최초 12개월 · 2028 기본/105℃ 검토안",
      "연간 환산 비교값은 확정 부족량·판매 가능량이 아님. 시나리오 가정은 04 시트 입력란을 대신하지 않음.")
legend(wsR, 3)
widths(wsR, {"A": 40, **{CL(i): 11 for i in range(2, 27)}, "AA": 40})
row = 5
wsR.cell(row=row, column=1, value="1. 기간 정의").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["구분", "시작", "종료", "정제기", "운전", "200 L 충진", "생산 조건", "상태"]); row += 1
for p_ in (("현재", dt.date(2026, 1, 1), f"={R('con_s')}-1", "1대", "단일 설비 순차", "이지켐·한솔 수동 (~'27.6)", "현재 조건 (03 선택 간격)", "제공"),
           ("공사", f"={R('con_s')}", f"={R('con_e')}", "1대", "가동 여부 미확인", "-", f'=IF(ISNUMBER({R("h_con")}),"가용 "&{R("h_con")}&" h/월","공사 중 가용시간 미확인")', "계획"),
           ("시운전", f"={R('trial_s')}", f"={R('rf_s')}-1", "1·2", "시운전", "-", f'=IF(ISNUMBER({R("b_trial")}),{R("b_trial")}&" Batch","양품·출하 승인 미확인")', "계획"),
           ("개선 적용 생산", f"={R('rf_s')}", "", "2대", "시간차 병행", "이지켐·한솔 ARS ('27.7~)", f'=IF(ISNUMBER({R("int_rf")}),{R("int_rf")}&" h/Batch","설비별 간격 미확인")', "계획"),
           ("최초 12개월", f"={R('rf_s')}", f"={R('f12_e')}", "2대", "", "", "47.2톤 연간 환산 기준", "제공"),
           ("2028 기본", dt.date(2028, 1, 1), dt.date(2028, 12, 31), "2대", "", "ARS", "47.2톤 기준 · 이지켐 월 840 kg 가정", "가정"),
           ("105℃ 검토안", f'=IF(ISNUMBER({R("t105")}),{R("t105")},"미정")', "", "2대", "103→105℃", "", "+3톤 · 50.2톤 산술", "검토안")):
    for j, v in enumerate(p_):
        style_cell(wsR.cell(row=row, column=1 + j, value=v), "link" if j in (1, 2) else "text", "yyyy-mm-dd" if j in (1, 2) else None)
    row += 1
row += 1
wsR.cell(row=row, column=1, value="2. Capa. 로드맵 (연간 환산)").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["구분", "연간 환산 (t)", "증가분 (t)", "월 환산 (t)", "Batch 상당/년", "설비별 등가 간격 (h)", "상태", "비고"]); row += 1
for lab, fx, add, nref, st, note in (("현재 (1대·103℃)", f"={R('capa_now')}", "", R('n_now'), "기준값", ""),
                                     ("참고: 단순 2배", f"={R('capa_x2')}", "", R('n_new'), "적용 안 함", "차이 5.2톤 확정 손실 아님"),
                                     ("개선 후 (2대)", f"={R('capa_rf')}", f"={R('capa_add')}", R('n_new'), "제공 (12개월 환산)", "ARS 미가산"),
                                     ("105℃ 적용 시", f"={R('capa_t')}", f"={R('capa_tadd')}", R('n_new'), "검토안", "미확정·시점 미정"),
                                     ("목표 (2028년 이후)", f"={R('target')}", "", R('n_new'), "목표", "50톤 수준")):
    vals = [lab, fx, add, f"=B{row}/12", f"=B{row}*1000/{R('kg_b')}", f"={nref}*{R('yh')}/E{row}", st, note]
    for j, v in enumerate(vals):
        style_cell(wsR.cell(row=row, column=1 + j, value=v), "key" if j == 1 else ("text" if j in (0, 6, 7) else "calc"), [None, "t", "t", "t", "b1", "h2", None, None][j])
    row += 1
row += 1
wsR.cell(row=row, column=1, value="3. 연도별 연간 환산 Capa. vs 출하 (산술 비교 — 실제 여유·부족 아님)").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["비교", "출하 (t)", "비교 Capa. (t)", "Capa. − 출하 (t)", "방향", "", "", "설명"]); row += 1
CMPY = {}


def SH(yr):
    return f"'01_Inputs'!$O${SHIP[(yr, 'tot')]}/1000"


h1 = f"SUM('01_Inputs'!C{SHIP[(2027, 'tot')]}:H{SHIP[(2027, 'tot')]})/1000"
h2 = f"SUM('01_Inputs'!I{SHIP[(2027, 'tot')]}:N{SHIP[(2027, 'tot')]})/1000"
for key, lab, sh, cap, note in (("26_now", "2026 제시분 vs 현재 26.2", f"={SH(2026)}", f"={R('capa_now')}", "CXMT·한솔 미포함"),
                                ("27_now", "2027 출하 vs 현재 26.2", f"={SH(2027)}", f"={R('capa_now')}", "17.2톤 초과"),
                                ("27_rf", "2027 출하 vs 개선 47.2 (연간 환산)", f"={SH(2027)}", f"={R('capa_rf')}", "3.8톤 — 2027 실제 여유 아님"),
                                ("27_h1", "2027 상반기 출하 vs 26.2 × 6/12 (참고)", f"={h1}", f"={R('capa_now')}*6/12", "공사·시운전 미반영 참고"),
                                ("27_h2", "2027 하반기 출하 vs 47.2 × 6/12 (균등 가동 참고)", f"={h2}", f"={R('capa_rf')}*6/12", "0.96톤 부족 방향"),
                                ("28_base", "2028 출하 가정 vs 47.2 (기본)", f"={SH(2028)}", f"={R('capa_rf')}", "기본 시나리오 — 방향은 E열"),
                                ("28_105", "2028 출하 가정 vs 50.2 (105℃ 12개월 적용 시, 검토안)", f"={SH(2028)}", f"={R('capa_t')}", "검토안 — 미확정")):
    vals = [lab, sh, cap, f"=C{row}-B{row}", f'=IF(D{row}<0,"부족 방향","여유 방향")', "", "", note]
    for j, v in enumerate(vals):
        style_cell(wsR.cell(row=row, column=1 + j, value=v), "key" if j == 3 else ("text" if j in (0, 4, 7) else "link"), "t" if j in (1, 2, 3) else None)
    CMPY[key] = row; row += 1
wsR.conditional_formatting.add(f"D{CMPY['26_now']}:D{CMPY['28_105']}", CellIsRule(operator="lessThan", formula=["0"], fill=red, font=redf))
row += 1
wsR.cell(row=row, column=1, value="4. 2027~2028 실제 생산 가능량 (04 시트 입력 기준 — 미확인 월은 부분합계)").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["구분"] + [f"{y % 100}.{m}" for y in (2027, 2028) for m in range(1, 13)] + ["비고"]); row += 1
for lab, col, nf in (("운전 조건", "D", None), ("생산 가능 Batch", "O", "int"), ("양품 생산 가능량 (kg)", "S", "kg0"), ("출하 (kg)", "Y", "kg0")):
    style_cell(wsR.cell(row=row, column=1, value=lab), "text", bold=True)
    for k in range(24):
        style_cell(wsR.cell(row=row, column=2 + k, value=f"='04_Monthly_2026_2028'!{col}{M0+12+k}"), "link", nf)
    row += 1
row += 1
wsR.cell(row=row, column=1, value="5. 2027~2028 공급 대응 시나리오 (참고 · 확정 아님) — 필요 선행재고 = 누적 (생산 − 출하) 최저점").font = font("EB002C", True, 11); row += 1
header(wsR, row, ["시나리오 가정", "S1 공사 중 중단", "S2 공사 중 현재 유지", "S3 50% · 안정화 3개월", "", "", "", "설명"]); row += 1
SP = {}
for key, lab, vals, nf, desc in (("con", "공사 기간 가동률 (현재 조건 대비)", [0, 1, 0.5], "pct", "2~5월"),
                                 ("trial", "시운전 월 양품 (kg)", [0, 0, 0], "kg0", "6월 — 출하 승인 미확인 → 0 가정"),
                                 ("stab", "개선 후 안정화 개월 (현재 수준 생산)", [0, 0, 3], "int", "7월부터"),
                                 ("rfkg", "개선 안정 후 월 생산 (kg)", [f"={R('capa_rf')}*1000/12"] * 3, "kg0", "47.2톤 ÷ 12 — 상한 참고"),
                                 ("t105", "2028 105℃ 12개월 적용 (1=적용)", [0, 0, 0], "int", "검토안 — 기본 0"),
                                 ("y", "양품률", [1, 1, 1], "pct", "100% = 상한")):
    style_cell(wsR.cell(row=row, column=1, value=lab), "text", bold=True)
    for j, v in enumerate(vals):
        style_cell(wsR.cell(row=row, column=2 + j, value=v), "est", nf)
    style_cell(wsR.cell(row=row, column=8, value=desc), "note"); SP[key] = row; row += 1
row += 1
header(wsR, row, ["월별 (kg)"] + [f"{y % 100}.{m}" for y in (2027, 2028) for m in range(1, 13)] + ["합계"]); row += 1
BASE = row
style_cell(wsR.cell(row=row, column=1, value="현재 조건 월 생산 가능량 (1대, Batch 상당×190)"), "text", bold=True)
for k in range(24):
    style_cell(wsR.cell(row=row, column=2 + k, value=f"='04_Monthly_2026_2028'!C{M0+12+k}*24/{C['basis_int']}*{R('kg_b')}"), "calc", "kg0")
row += 1
SHIPR = row
style_cell(wsR.cell(row=row, column=1, value="출하 (4개 고객)"), "text", bold=True)
for k in range(24):
    style_cell(wsR.cell(row=row, column=2 + k, value=f"='04_Monthly_2026_2028'!Y{M0+12+k}"), "link", "kg0")
style_cell(wsR.cell(row=row, column=26, value=f"=SUM(B{row}:Y{row})"), "key", "kg0"); row += 1
SCN = {}
for j, s in enumerate(["S1", "S2", "S3"]):
    pc = CL(2 + j); pr = row
    style_cell(wsR.cell(row=row, column=1, value=f"{s} 생산 (kg)"), "text", bold=True)
    for k in range(24):
        col = CL(2 + k); mm = k + 1
        fx = (f"=IF({mm}=1,{col}{BASE},IF({mm}<=5,{col}{BASE}*${pc}${SP['con']},IF({mm}=6,${pc}${SP['trial']},"
              f"IF({mm}-6<=${pc}${SP['stab']},{col}{BASE},IF(AND({mm}>12,${pc}${SP['t105']}=1),{R('capa_t')}*1000/12,${pc}${SP['rfkg']})))))*${pc}${SP['y']}")
        style_cell(wsR.cell(row=row, column=2 + k, value=fx), "calc", "kg0")
    style_cell(wsR.cell(row=row, column=26, value=f"=SUM(B{row}:Y{row})"), "key", "kg0"); row += 1
    cr = row
    style_cell(wsR.cell(row=row, column=1, value=f"{s} 누적 (생산 − 출하)"), "text")
    for k in range(24):
        col = CL(2 + k)
        style_cell(wsR.cell(row=row, column=2 + k, value=f"={col}{pr}-{col}{SHIPR}" if k == 0 else f"={CL(1+k)}{cr}+{col}{pr}-{col}{SHIPR}"), "calc", "kg0")
    row += 1
    SCN[s] = (pr, cr)
wsR.conditional_formatting.add(f"B{SCN['S1'][1]}:Y{SCN['S3'][1]}", CellIsRule(operator="lessThan", formula=["0"], fill=red, font=redf))
row += 1
header(wsR, row, ["시나리오 결과", "S1", "S2", "S3", "", "", "", "해석"]); row += 1
RES = {}
for lab, fn, nf, desc in (("2027 생산 (kg)", lambda s: f"=SUM(B{SCN[s][0]}:M{SCN[s][0]})", "kg0", "참고 — 실제 아님"),
                          ("2027 생산 − 출하 (kg)", lambda s: f"=SUM(B{SCN[s][0]}:M{SCN[s][0]})-SUM(B{SHIPR}:M{SHIPR})", "kg0", ""),
                          ("2028 생산 − 출하 (kg)", lambda s: f"=SUM(N{SCN[s][0]}:Y{SCN[s][0]})-SUM(N{SHIPR}:Y{SHIPR})", "kg0", "47.2 균등 환산 기준"),
                          ("필요 선행재고 2027-01-01 (kg)", lambda s: f"=MAX(0,-MIN(B{SCN[s][1]}:Y{SCN[s][1]}))", "kg0", "24개월 결품 없이 대응할 최소 기초재고"),
                          ("누적 최저 시점", lambda s: f'=IF(MIN(B{SCN[s][1]}:Y{SCN[s][1]})<0,INDEX($B${BASE-1}:$Y${BASE-1},MATCH(MIN(B{SCN[s][1]}:Y{SCN[s][1]}),B{SCN[s][1]}:Y{SCN[s][1]},0)),"-")', None, "연.월")):
    style_cell(wsR.cell(row=row, column=1, value=lab), "text", bold=True)
    for j, s in enumerate(["S1", "S2", "S3"]):
        style_cell(wsR.cell(row=row, column=2 + j, value=fn(s)), "key", nf)
    style_cell(wsR.cell(row=row, column=8, value=desc), "note"); RES[lab] = row; row += 1
row += 1
wsR.cell(row=row, column=1, value="6. 105℃ 검토안 확인사항 (미확정 · 적용 시점 미정 · 검증 및 승인 필요)").font = font("EB002C", True, 11); row += 1
for t in ["열 안정성", "Dimer 형성", "Unknown impurity 증가", "Yield 영향", "고객 승인", "적용 시점 — 확정 시 01_Inputs '105℃ 적용 시작 월' 입력 → 04 시트 월별 반영"]:
    style_cell(wsR.cell(row=row, column=1, value="· " + t), "text"); row += 1

# ======================================================== 06_Report_Summary
wsS = wb.create_sheet("06_Report_Summary")
title(wsS, "06_Report_Summary | PPT 사용 수치 · 표 · 차트 (수식 연동)", "PPT 수치는 이 시트와 일치해야 함.")
widths(wsS, {"A": 44, "B": 14, "C": 12, "D": 10, **{CL(i): 11 for i in range(5, 16)}})
row = 4
wsS.cell(row=row, column=1, value="1. 핵심 수치").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["항목", "값", "단위", "시트"]); row += 1
KPI = {}
Mq = "'04_Monthly_2026_2028'!"


def cy(k):
    return f"'05_Reflux_Scenarios'!$D${CMPY[k]}"


kp = [("t57c", f"={R('t57c')}", "h", "01", "h"), ("t_ref", f"={R('t_ref')}", "h", "01", "h"),
      ("lt_5g", f"={C['lt_5g']}", "h", "03", "h"), ("lt_ez", f"={C['lt_ez']}", "h", "03", "h"), ("el_5g", f"={C['el_5g']}", "근무일", "03", "h2"),
      ("occ", f"={C['occ']}", "h", "03", "h"), ("th_bm", f"={C['th_bm']}", "Batch", "03", "b"), ("th_kgm", f"={C['th_kgm']}", "kg", "03", "kg0"),
      ("mx_ref", f"={C['mx_ref']}", "h", "03", "h"), ("mx_gap", f"={C['mx_gap']}", "h", "03", "h"), ("mx_kgm", f"={C['mx_kgm']}", "kg", "03", "kg0"),
      ("mx_kgy12", f"={C['mx_kgy12']}", "kg", "03", "kg0"), ("mx_int", f"={C['mx_int']}", "h", "03", "h2"),
      ("ob_int", f"={C['ob_int']}", "h", "03", "h2"), ("ob_lo", f"={C['ob_lo']}", "h", "03", "h2"), ("ob_hi", f"={C['ob_hi']}", "h", "03", "h2"),
      ("ob_med", f"={C['ob_med']}", "h", "03", "h"), ("ob_n", f"={C['ob_n']}", "", "03", "int"), ("ob_days", f"={C['ob_days']}", "일", "03", "int"),
      ("ob_n_dates", f"={C['ob_n_dates']}", "", "03", "int"), ("ob_vs47", f"={C['ob_vs47']}", "h", "03", "h2"), ("ob_vs45", f"={C['ob_vs45']}", "h", "03", "h2"),
      ("ob_b30", f"={C['ob_b30']}", "Batch", "03", "b"), ("ob_b31", f"={C['ob_b31']}", "Batch", "03", "b"), ("ob_kgm", f"={C['ob_kgm']}", "kg", "03", "kg0"),
      ("ob_kgy", f"={C['ob_kgy']}", "kg", "03", "kg0"), ("ob_share3", f"={C['ob_share3']}", "%", "03", "pct"), ("ob_dense", f"={C['ob_dense']}", "", "02", "int"),
      ("chk_diff", f"={CHK[0]}", "건", "02", "int"), ("cp_b", f"={C['cp_b']}", "Batch", "03", "0.000"), ("cp_kground", f"={C['cp_kground']}", "kg", "03", "kg0"),
      ("cp_int", f"={C['cp_int']}", "h", "03", "h2"), ("cp_bm", f"={C['cp_bm']}", "Batch", "03", "b"),
      ("x2_gap", f"={C['x2_gap']}", "t", "03", "t"), ("rf_int_each", f"={C['rf_int_each']}", "h", "03", "h2"), ("rf_int_all", f"={C['rf_int_all']}", "h", "03", "h2"),
      ("rf_int_diff", f"={C['rf_int_diff']}", "h", "03", "h2"), ("off", f"={C['off']}", "h", "03", "h"), ("g_int", f"={C['g_int']}", "h", "03", "h"),
      ("fill_gap_min", f"={C['fill_gap_min']}", "h", "03", "h"),
      ("tk1_asis", f"='03_Capacity_Model'!$E${C['tk0']}", "h", "03", "h"), ("tk1_tobe", f"='03_Capacity_Model'!$F${C['tk0']}", "h", "03", "h"),
      ("tk2_tobe", f"='03_Capacity_Model'!$F${C['tk0']+1}", "h", "03", "h"), ("tk1_occ", f"='03_Capacity_Model'!$D${C['tk0']}", "h", "03", "h"),
      ("gb27", f"={C['gb27']}", "h", "03", "h"), ("gb28", f"={C['gb28']}", "h", "03", "h"), ("gbmax", f"={C['gbmax']}", "h", "03", "h"),
      ("gb_ld1", f"='03_Capacity_Model'!$E${C['pp0']+1}", "%", "03", "pct"), ("gb_ld2", f"='03_Capacity_Model'!$E${C['pp0']+2}", "%", "03", "pct"),
      ("c200_27", f"={C['c200_27']}", "용기", "03", "int"), ("c200_28", f"={C['c200_28']}", "용기", "03", "int"),
      ("sh26", f"={SH(2026)}", "t", "01", "t"), ("sh27", f"={SH(2027)}", "t", "01", "t"), ("sh28", f"={SH(2028)}", "t", "01", "t"),
      ("sh27h1", f"={h1}", "t", "01", "t"), ("sh27h2", f"={h2}", "t", "01", "t"),
      ("d26_now", f"={cy('26_now')}", "t", "05", "t"), ("d27_now", f"={cy('27_now')}", "t", "05", "t"), ("d27_rf", f"={cy('27_rf')}", "t", "05", "t"),
      ("d27_h1", f"={cy('27_h1')}", "t", "05", "t"), ("d27_h2", f"={cy('27_h2')}", "t", "05", "t"), ("d28_base", f"={cy('28_base')}", "t", "05", "t"),
      ("d28_105", f"={cy('28_105')}", "t", "05", "t"), ("cap27h2", f"='05_Reflux_Scenarios'!$C${CMPY['27_h2']}", "t", "05", "t"),
      ("cap27h1", f"='05_Reflux_Scenarios'!$C${CMPY['27_h1']}", "t", "05", "t"),
      ("p26", f"={Mq}$S${SR[2026]}", "kg", "04", "kg0"), ("pb26", f"={Mq}$O${SR[2026]}", "Batch", "04", "int"),
      ("p27", f"={Mq}$S${SR[2027]}", "kg", "04", None), ("p28", f"={Mq}$S${SR[2028]}", "kg", "04", None),
      ("p27_1", f"={Mq}$S${M0+12}", "kg", "04", "kg0"), ("pb27_1", f"={Mq}$O${M0+12}", "Batch", "04", "int"),
      ("basis_int", f"={C['basis_int']}", "h", "03", "h2"),
      ("nb27", f"={Mq}$Z${SR[2027]}", "Batch", "04", "b1"), ("nb28", f"={Mq}$Z${SR[2028]}", "Batch", "04", "b1"),
      ("ez27", f"={Mq}$W${SR[2027]}", "kg", "04", "kg0"), ("hs27", f"={Mq}$X${SR[2027]}", "kg", "04", "kg0"), ("cx27", f"={Mq}$V${SR[2027]}", "kg", "04", "kg0"),
      ("ez28", f"={Mq}$W${SR[2028]}", "kg", "04", "kg0"), ("cx28", f"={Mq}$V${SR[2028]}", "kg", "04", "kg0"), ("hs28", f"={Mq}$X${SR[2028]}", "kg", "04", "kg0"),
      ("ak27", f"={Mq}$AK${SR[2027]}", "h", "04", "h"), ("ak28", f"={Mq}$AK${SR[2028]}", "h", "04", "h"),
      ("ah27", f"={Mq}$AH${SR[2027]}", "h", "04", "h"), ("al27", f"={Mq}$AL${SR[2027]}", "h", "04", "h"), ("ac27", f"={Mq}$AC${SR[2027]}", "h", "04", "h"),
      ("al28", f"={Mq}$AL${SR[2028]}", "h", "04", "h"), ("ac28", f"={Mq}$AC${SR[2028]}", "h", "04", "h")]
for s in ("S1", "S2", "S3"):
    j = "BCD"[["S1", "S2", "S3"].index(s)]
    for lab, key in (("2027 생산 (kg)", "p"), ("2027 생산 − 출하 (kg)", "d27"), ("2028 생산 − 출하 (kg)", "d28"),
                     ("필요 선행재고 2027-01-01 (kg)", "inv"), ("누적 최저 시점", "low")):
        kp.append((f"{s}_{key}", f"='05_Reflux_Scenarios'!${j}${RES[lab]}", "kg", "05", None if key == "low" else "kg0"))
for k, fx, unit, src, nf in kp:
    style_cell(wsS.cell(row=row, column=1, value=k), "text"); style_cell(wsS.cell(row=row, column=2, value=fx), "link", nf)
    style_cell(wsS.cell(row=row, column=3, value=unit), "text"); style_cell(wsS.cell(row=row, column=4, value=src), "text")
    KPI[k] = row; row += 1
row += 1
wsS.cell(row=row, column=1, value="2. 생산능력 비교 (03 시트 7절)").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["구분", "간격 (h)", "월 Batch", "월 생산 kg", "연 환산 t", "26.2 대비"]); row += 1
CMPS = row
for k in range(4):
    cr_ = CMP0 + k
    for j, col in enumerate("ABCDEF"):
        style_cell(wsS.cell(row=row, column=1 + j, value=f"='03_Capacity_Model'!{col}{cr_}"), "link", [None, "h2", "b", "kg0", "t", "t"][j])
    row += 1
ch = BarChart(); ch.type = "bar"; ch.title = "연 환산 비교 (t)"; ch.style = 10; ch.legend = None
ch.add_data(Reference(wsS, min_col=5, min_row=CMPS - 1, max_row=CMPS + 3), titles_from_data=True)
ch.set_categories(Reference(wsS, min_col=1, min_row=CMPS, max_row=CMPS + 3)); ch.height = 6; ch.width = 14
wsS.add_chart(ch, f"H{CMPS-1}")
row += 6
wsS.cell(row=row, column=1, value="3. 36개월 출하 · 참고 Capa. · 생산 가능량 (kg)").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["연월", "운전 조건", "하이닉스", "CXMT", "이지켐", "한솔", "출하 합계", "참고 Capa. 월 환산", "참고 Capa.−출하",
                  "생산 가능 Batch", "양품 생산 가능량", "5 Gal 충진 h", "200 L 용기", "계산 상태"]); row += 1
MS0 = row
for k in range(36):
    rr = M0 + k
    for j, col in enumerate(["A", "D", "U", "V", "W", "X", "Y", "T", "AO", "O", "S", "AB", None, "BD"]):
        fx = f"={Mq}AD{rr}+{Mq}AE{rr}" if col is None else f"={Mq}{col}{rr}"
        style_cell(wsS.cell(row=row, column=1 + j, value=fx), "link", ["ym", None, "kg", "kg", "kg", "kg", "kg", "kg", "kg", "int", "kg0", "h", "b1", None][j])
    row += 1
wsS.conditional_formatting.add(f"I{MS0}:I{MS0+35}", CellIsRule(operator="lessThan", formula=["0"], fill=red, font=redf))
ch2 = BarChart(); ch2.title = "월별 출하 합계 vs 참고 Capa. 월 환산 (kg)"; ch2.style = 10
ch2.add_data(Reference(wsS, min_col=7, min_row=MS0 - 1, max_row=MS0 + 35), titles_from_data=True)
ch2.add_data(Reference(wsS, min_col=8, min_row=MS0 - 1, max_row=MS0 + 35), titles_from_data=True)
ch2.set_categories(Reference(wsS, min_col=1, min_row=MS0, max_row=MS0 + 35)); ch2.height = 8; ch2.width = 28
wsS.add_chart(ch2, f"P{MS0}")
row += 1
wsS.cell(row=row, column=1, value="4. 과거 간격 분포 (#39~#92)").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["간격", "건수"]); row += 1
GD0 = row
for j, (lab, col) in enumerate((("1일", "O"), ("2일", "P"), ("3일", "Q"), ("4일", "R"), ("5일", "S"))):
    style_cell(wsS.cell(row=row, column=1, value=lab), "text"); style_cell(wsS.cell(row=row, column=2, value=f"='02_Batch_Raw'!{col}{rt}"), "link", "int"); row += 1
row += 1
wsS.cell(row=row, column=1, value="5. 주요 미확인 사항").font = font("EB002C", True, 11); row += 1
header(wsS, row, ["확인사항", "담당", "영향"]); row += 1
for t, who, eff in (("표시일자 의미 · 실제 시각 · 계획 대비 실적 · Total 46 vs 번호 54 · 색상 의미", "생산", "관측 간격 67.5 h 전제"),
                    ("이송 중 정제기 점유 · 세척·전환 · 대기 · 보수 시간", "생산", "점유 47~53 h, 비점유 약 20 h 구성"),
                    ("정제기 2 기동 시차 · 설비별 간격 · Mix·리사이클 준비 시간", "생산·설비", "47.2톤 달성 조건 (설비별 등가 70.5 h)"),
                    ("공사 중 가동 · 시운전 양품·출하 승인 · 초기 안정화", "설비·품질", "2027 실제 생산량"),
                    ("190 kg vs 9병 180 kg 잔량 10 kg 처리 · 200 L 후 잔량 5 Gal 충진", "생산·품질", "충진량·재고"),
                    ("총 투입·재순환·양품 실측 → 수율 (구두 160~220 kg 범위)", "생산", "Batch량 검증"),
                    ("고객별 합격률 · 이지켐 색도", "품질", "출하 가능 양품"),
                    ("한솔 수동 충진시간 · 이지켐·한솔 ARS 충진시간 ('27.7~)", "생산·설비", "200 L 작업량"),
                    ("글로브 박스 교대 · Tank 수·용량 · QC 근무시간", "생산", "5 Gal 267 h/월 · Tank 회전"),
                    ("기초재고 · 목표재고 · 고객 간 우선순위", "생산관리·영업", "월별 재고·배정"),
                    ("2026 CXMT·한솔 물량 · 하이닉스 실적/계획 구분 · 2028 수요 확정", "영업", "출하 합계"),
                    ("105℃: 열 안정성·Dimer·Unknown impurity·Yield · 승인 · 적용 시점", "품질·기술", "50.2톤 검토안")):
    for j, v in enumerate((t, who, eff)):
        style_cell(wsS.cell(row=row, column=1 + j, value=v), "text")
    row += 1
wsS.freeze_panes = "A4"

C["lt0"] = lt0
add_ppt_sheets(wb, R, C, {"M0": M0, "SR": SR, "B0": B0, "BL": BL, "SUMROW": SUMROW, "CMP0": CMP0, "CMPY": CMPY}, L)
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
    tc = {"01_Inputs": "2E75B6", "02_Batch_Raw": "7F7F7F", "03_Capacity_Model": "EB002C",
          "04_Monthly_2026_2028": "FF7900", "05_Reflux_Scenarios": "70AD47", "06_Report_Summary": "404040"}.get(ws.title)
    if tc:
        ws.sheet_properties.tabColor = tc
wb.save(OUT)
json.dump({"KPI": KPI, "M0": M0, "SR": SR, "MX": MX, "SUMROW": SUMROW, "SHIP": {str(k): v for k, v in SHIP.items()}, "CMP0": CMP0,
           "MS0": MS0, "GD0": GD0, "G0": C["g0"], "GN": C["gn"], "RES": RES, "CMPY": CMPY}, open(OUT + ".map.json", "w"), ensure_ascii=False)
print("saved", OUT)
