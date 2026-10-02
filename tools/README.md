# CpZr 생산능력 자료 생성 절차 (v3)

1. `python build_xlsx.py raw.xlsx` — 계산 Excel 생성 (모든 결과는 수식, `01_입력`의 이름 정의를 참조).
2. 사본을 LibreOffice로 재계산 (`recalc.py calc.xlsx`)해 수식 오류 0 확인.
3. `python inject_cache.py raw.xlsx calc.xlsx SKTC_CpZr_Capacity_RawData_Calculation_2026_2028.xlsx` — 원본 수식에 계산값만 넣어 배포본 생성.
   (LibreOffice가 다시 저장한 파일은 '01_입력'! 등 숫자 시작 시트명의 따옴표가 빠져 Excel 복구 오류가 나므로 배포 금지.)
4. `python build_pptx.py template/base.pptx calc.xlsx raw.xlsx.map.json SKTC_CpZr_Capacity_Reflux_Roadmap_2027_2028.pptx`
   — SK trichem 양식의 공정도·머리글을 재사용하고 표·마일스톤·달력 표는 Excel 값으로 작성 (12장, 그래프 없음).

## Excel 시트
| 시트 | 내용 |
|---|---|
| 01_입력 | 모든 기준값 · 월별 기존 생산계획 · 고객별 출하계획 (이 시트만 수정) |
| 02_공정시간 | 57 h 구성 · 경로별 단순 합계 · 시간 구분 |
| 03_Batch주기 | 생산계획 완료 주기 · 순수 작업 vs 대기 · 평균 마일스톤 · 정제기 2대 운영 예시 |
| 04_Capa기준 | 연간 → 월 기준 (대정비 반영 ÷11) · 일정 · 2026 출하 |
| 05_월별_Capa기준 / 06_월별_생산계획 | 두 버전의 2027·2028 월별 생산 vs 출하 |
| 07_버전비교 | 연간 비교 · 24개월 누적 · 차트 |
| 08_후공정부하 | 5 Gal · 200 L 충진 · OQC · 검사 작업량 |
| 09_확인사항 | 생산팀 확인 내용 · 조건별 생산량 입력란 |
| 10_과거Batch_참고 | 2026 생산계획 표시일자 원자료 (참고) |

## PPT 구성
1 As-is 공정·마일스톤 · 2 10월 생산계획 완료 주기 · 3 순수 작업 vs 대기 · 4 To-be 공정·일정 · 5 To-be 운영 마일스톤 · 6 버전 비교 ·
7~10 2027/2028 × (Capa. 기준 / 기존 생산계획) · 11 후공정 부하 · 12 생산팀 확인 사항
