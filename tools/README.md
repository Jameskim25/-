# CpZr 생산능력 자료 생성 절차 (v4 · Case I~III)

1. `python build_xlsx.py raw.xlsx` — 계산 Excel 생성 (모든 결과는 수식, `01_입력`의 이름 정의를 참조).
2. 사본을 LibreOffice로 재계산 (`recalc.py calc.xlsx`)해 수식 오류 0 확인.
3. `python inject_cache.py raw.xlsx calc.xlsx SKTC_CpZr_Capacity_RawData_Calculation_2026_2028.xlsx` — 원본 수식에 계산값만 넣어 배포본 생성.
   (LibreOffice가 다시 저장한 파일은 '01_입력'! 등 숫자 시작 시트명의 따옴표가 빠져 Excel 복구 오류가 나므로 배포 금지.)
4. `python build_pptx.py template/base.pptx calc.xlsx raw.xlsx.map.json SKTC_CpZr_Capacity_Reflux_Roadmap_2027_2028.pptx`
   — SK trichem 양식의 공정도·머리글을 재사용하고 표·마일스톤·달력 표는 Excel 값으로 작성 (12장, 그래프 없음).

## Excel 시트
| 시트 | 내용 |
|---|---|
| 01_입력 | 모든 기준값 (Batch size 200 · Yield 99% → 198 kg/Batch, 월 Batch 12/20/6) · 10월 계획 완료일 (이 시트만 수정) |
| 02_공정시간 | 57 h 구성 · 경로별 단순 합계 · 시간 구분 |
| 03_Batch주기 | 10월 생산계획 완료 주기 · 순수 작업 vs 대기 · 평균 마일스톤 · 설비별 주기 비교 · 2대 운영 예시 |
| 04_Capa기준 | 월 Batch 기준 생산능력 · Case별 리플럭스 도입 일정 |
| 05_Case1 / 06_Case2 / 07_Case3 | '26.9~'27.12 월별 Batch·생산·판매·고객·재고·재고일 (협의 자료) |
| 08_Case비교 | 2027 연간 비교 · 월별 재고/재고일 |
| 09_후공정부하 | 5 Gal · 200 L 충진 · OQC · 검사 작업량 (Case II 판매 기준) |
| 10_확인사항 | 생산팀 확인 내용 · 협의 사항 · 조건별 생산량 입력란 |
| 11_과거Batch_참고 | 2026 생산계획 표시일자 원자료 (참고) |

## PPT 구성 (12장, 그래프 없음)
1 As-is 공정·마일스톤 · 2 10월 생산계획 완료 주기 · 3 순수 작업 vs 대기 · 4 To-be 공정·Case별 일정 · 5 To-be 운영 마일스톤 ·
6 Case 비교 · 7~9 Case I / II / III · 10 후공정 부하 · 11 생산팀 협의 사항 · 12 생산팀 확인 사항
