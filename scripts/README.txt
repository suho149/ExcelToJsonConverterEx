[실행 방법]
1) CreateInsertSql.exe 실행
2) 기본 입력 폴더: excel\

[동작]
- excel\ 폴더에서 엑셀 파일(xlsx/xlsm/xls) 1개를 자동 인식해 읽고, Sheet4에 JSON 변환 결과를 작성합니다.
- exceldata.xlsx 파일도 기존과 동일하게 우선 인식합니다.
- Sheet4 기준으로 INSERT SQL을 생성합니다.
- SQL 파일은 output\sheet4_insert_YYYYMMDD_HHMMSS.sql 로 저장됩니다.

[주의]
- excel\ 폴더에 엑셀 파일이 2개 이상이면 어떤 파일을 쓸지 결정할 수 없어 실행이 중단됩니다.
- 실행 중에는 대상 엑셀 파일을 Excel에서 열어두면 안 됩니다.
