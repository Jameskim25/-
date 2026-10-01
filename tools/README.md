# CpZr 생산능력 자료 생성 절차 (v2)

1. `build_xlsx.py <원본 PPT> <수정본 PPT> <out.xlsx>` — 계산 Excel 생성 (모든 결과는 수식).
2. LibreOffice로 재계산 (`recalc.py out.xlsx`) — 캐시 값 생성, 수식 오류 0 확인.
3. PPT 구조 정리: 수정본 PPT에서 2027장(slide59)을 복제해 2028장 생성, 2026장(slide58)은 3장에 통합하므로 `sldIdLst`에서 제거 후 정리.
4. `edit_pptx.py <구조 정리본> <수정본 PPT> <out.xlsx> <out.pptx>` — 기존 SK trichem 양식 도형·표를 재사용해 텍스트·막대·표를 수정하고 수치는 Excel에서 읽음.

## PPT 붙여넣기용 가로형 시트 (P1~P5)
`ppt_sheets.py`가 생성. 각 시트는 PPT 장과 1:1 (P1 공정시간, P2 일정·30일 Gantt, P3 과거 Batch·2026, P4 2027 월별, P5 2028 월별).
모든 칸은 01_Inputs·02~05 시트와 수식 연결 — 입력을 바꾸면 갱신. 열 폭·글꼴(Noto Sans KR)·색은 PPT 표 크기에 맞춤.
표 범위를 복사 → PPT에서 "원본 서식 유지" 또는 "그림"으로 붙여넣기.
