# SKTC 중국법인 설립 검토 (강소성 우선) — 2026-10-06

산출물
- SKTC_중국법인설립_강소성우선_입지및인허가검토_20261006.pptx — SK trichem 양식(tools/template/base.pptx) 적용, 본문 28장 + 부록 13장
- SKTC_중국법인설립_Raw_File_Data_20261006.xlsx — 19개 시트, 수식 321개(재계산 오류 0)
- SKTC_중국법인설립_강소성우선_입지및인허가검토_20261006.pdf — LibreOffice 렌더링 미리보기

생성 절차
1. `python build_cn_xlsx.py raw.xlsx` (데이터: cn_data.py)
2. LibreOffice 재계산 사본 calc.xlsx (`recalc.py`)
3. `python ../tools/inject_cache.py raw.xlsx calc.xlsx <배포 xlsx>`
4. `python build_cn_pptx.py ../tools/template/base.pptx calc.xlsx <china provinces geojson> <배포 pptx>`
   (지도 경계: github.com/longwosion/geojson-map-china china.json, 단순화)

조사 한계: 작업 환경 네트워크 정책으로 중국 정부 사이트 원문을 직접 열지 못해 검색 결과 요약으로 확인했다.
공식 출처라도 '공식 출처·요약 확인'으로 표시했으며 원문 조문 대조가 필요하다.
