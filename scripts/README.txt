[실행 방법]
1) ExcelToJsonConverter.exe 실행
2) 기본 입력 파일: excel\exceldata.xlsx

[동작]
- excel\exceldata.xlsx를 읽어 Sheet4에 JSON 변환 결과를 작성합니다.
- Sheet4 기준으로 INSERT SQL을 생성합니다.
- SQL 파일은 output\sheet4_insert_YYYYMMDD_HHMMSS.sql 로 저장됩니다.

[주의]
- 파일명은 반드시 exceldata.xlsx 여야 합니다.
- 실행 중에는 exceldata.xlsx를 Excel에서 열어두면 안 됩니다.
