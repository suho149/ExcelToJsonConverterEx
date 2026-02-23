package demo.tojson;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.AccessDeniedException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelToJsonConverter {

    // classpath 기준 엑셀 파일 위치 (src/main/resources 아래)
    // src/main/resources/excel/exceldata.xlsx
    private static final String EXCEL_RESOURCE_PATH = "excel/exceldata.xlsx";
    private static final Path PROJECT_DEFAULT_EXCEL_PATH = Paths.get("src/main/resources").resolve(EXCEL_RESOURCE_PATH);
    private static final Path APP_DEFAULT_EXCEL_PATH = Paths.get("excel").resolve("exceldata.xlsx");

    // 시트 이름 (엑셀에서 실제 이름 그대로 사용)
    private static final String SHEET1_NAME = "Sheet1";
    private static final String SHEET3_NAME = "Sheet3";
    private static final String SHEET4_NAME = "Sheet4";

    // Sheet1 기준
    private static final int SHEET1_HEADER_ROW_INDEX = 1;      // 엑셀 2행 (헤더: company_id, ...)
    private static final int SHEET1_DATA_START_ROW_INDEX = 2;  // 엑셀 3행부터 데이터

    // Sheet3 기준
    private static final int SHEET3_START_ROW_INDEX = 2;       // 엑셀 3행부터 매핑
    private static final int SHEET3_SOURCE_COL_INDEX = 2;      // C열: 통합정보시스템 컬럼명
    private static final int SHEET3_JSONKEY_COL_INDEX = 3;     // D열: JSON 키 이름

    // Sheet4 기준 (출력 시트)
    private static final int SHEET4_HEADER_ROW_INDEX = 0;         // 헤더는 1행
    private static final int SHEET4_DATA_START_ROW_INDEX = 1;     // 데이터는 2행부터
    private static final int SHEET4_COMPANY_IDX_COL_INDEX = 0;    // A열
    private static final int SHEET4_JSON_COL_INDEX = 1;           // B열

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // CLI 환경(서버/터미널)에서도 POI 컬럼 폭 계산이 안정적으로 동작하도록 headless 고정
        System.setProperty("java.awt.headless", "true");

        Path inputExcelPath = resolveDefaultExcelPath().toAbsolutePath().normalize();
        Path outputExcelPath = inputExcelPath;

        if (!Files.exists(inputExcelPath)) {
            throw new IllegalStateException("입력 엑셀 파일을 찾을 수 없습니다: " + inputExcelPath
                    + System.lineSeparator()
                    + "확인 경로: " + PROJECT_DEFAULT_EXCEL_PATH.toAbsolutePath().normalize()
                    + " 또는 " + APP_DEFAULT_EXCEL_PATH.toAbsolutePath().normalize());
        }

        // Windows에서는 원본 파일을 File로 열어 둔 상태에서 교체(move)하면 잠금 오류가 날 수 있어
        // InputStream으로 읽어 메모리로 로드한 뒤 저장 시 교체한다.
        Workbook workbook;
        try (InputStream is = Files.newInputStream(inputExcelPath)) {
            workbook = WorkbookFactory.create(is);
        }

        try (workbook) {

            Sheet sheet1 = workbook.getSheet(SHEET1_NAME);
            Sheet sheet3 = workbook.getSheet(SHEET3_NAME);
            Sheet sheet4 = workbook.getSheet(SHEET4_NAME);

            if (sheet1 == null || sheet3 == null) {
                throw new IllegalStateException("Sheet1 또는 Sheet3를 찾을 수 없습니다.");
            }

            if (sheet4 == null) {
                sheet4 = workbook.createSheet(SHEET4_NAME);
            }
            clearSheet(sheet4);
            writeSheet4Header(sheet4);

            // Sheet3에서 매핑 정보 읽기 (C열 = 원본 컬럼명, D열 = JSON 키)
            LinkedHashMap<String, String> fieldMappings = loadFieldMappings(sheet3);

            // Sheet1 헤더(2행) 읽어서 "컬럼명 → 인덱스" 맵 만들기
            Map<String, Integer> headerIndexMap = buildHeaderIndexMap(sheet1);

            Integer companyIdColIndex = headerIndexMap.get("company_id");
            if (companyIdColIndex == null) {
                throw new IllegalStateException("Sheet1 헤더에 'company_id' 컬럼이 없습니다.");
            }

            // 데이터 행을 돌면서 JSON 생성 후 Sheet4에 기록
            ObjectMapper mapper = new ObjectMapper();
            int outputRowIndex = SHEET4_DATA_START_ROW_INDEX;

            int lastRow = sheet1.getLastRowNum();
            for (int rowIndex = SHEET1_DATA_START_ROW_INDEX; rowIndex <= lastRow; rowIndex++) {
                Row row = sheet1.getRow(rowIndex);
                if (row == null) {
                    continue;
                }

                String companyId = getCellString(row.getCell(companyIdColIndex));
                if (companyId == null || companyId.isBlank()) {
                    // company_id 비어 있으면 이후 행은 없다고 보고 종료
                    break;
                }

                // contents 객체 만들기
                ObjectNode contentsNode = mapper.createObjectNode();

                for (Map.Entry<String, String> entry : fieldMappings.entrySet()) {
                    String sourceColumnName = entry.getKey();  // 예: company_join_date, 없음, 확인중
                    String jsonKey = entry.getValue();         // 예: defenseDesignationDate

                    String value;

                    if (sourceColumnName == null ||
                            sourceColumnName.isBlank() ||
                            "없음".equals(sourceColumnName) ||
                            "확인중".equals(sourceColumnName)) {
                        // 원본 컬럼이 없거나 "없음"/"확인중"인 경우 → 빈 문자열
                        value = "";
                    } else {
                        Integer colIndex = headerIndexMap.get(sourceColumnName);
                        if (colIndex == null) {
                            // 매핑에는 있는데 실제 Sheet1 헤더엔 없으면 빈값
                            value = "";
                        } else {
                            value = getCellString(row.getCell(colIndex));
                        }
                    }

                    contentsNode.put(jsonKey, value == null ? "" : value);
                }

                // JSON 전체 구조 만들기
                ObjectNode companyInfoNode = mapper.createObjectNode();
                companyInfoNode.set("contents", contentsNode);

                ObjectNode dataNode = mapper.createObjectNode();
                dataNode.set("companyInfo", companyInfoNode);

                ObjectNode rootNode = mapper.createObjectNode();
                rootNode.set("data", dataNode);

                String jsonString = mapper.writeValueAsString(rootNode);

                Row outputRow = sheet4.createRow(outputRowIndex++);
                outputRow.createCell(SHEET4_COMPANY_IDX_COL_INDEX).setCellValue(companyId);
                outputRow.createCell(SHEET4_JSON_COL_INDEX).setCellValue(jsonString);
            }

            sheet4.autoSizeColumn(SHEET4_COMPANY_IDX_COL_INDEX);

            saveWorkbookSafely(workbook, outputExcelPath);

            System.out.println("변환 완료: " + outputExcelPath + " (" + SHEET4_NAME + " 시트)");
        }
    }

    private static Path resolveDefaultExcelPath() {
        if (Files.exists(PROJECT_DEFAULT_EXCEL_PATH)) {
            return PROJECT_DEFAULT_EXCEL_PATH;
        }
        if (Files.exists(APP_DEFAULT_EXCEL_PATH)) {
            return APP_DEFAULT_EXCEL_PATH;
        }
        return PROJECT_DEFAULT_EXCEL_PATH;
    }

    /**
     * Sheet3에서 C열(통합정보시스템 컬럼명)과 D열(JSON 키)을 읽어서
     * "원본컬럼명 → jsonKey" 매핑을 만든다.
     * LinkedHashMap을 쓰는 이유는 순서를 유지하기 위해서.
     */
    private static LinkedHashMap<String, String> loadFieldMappings(Sheet mappingSheet) {
        LinkedHashMap<String, String> mappings = new LinkedHashMap<>();

        int lastRow = mappingSheet.getLastRowNum();
        for (int rowIndex = SHEET3_START_ROW_INDEX; rowIndex <= lastRow; rowIndex++) {
            Row row = mappingSheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            Cell srcCell = row.getCell(SHEET3_SOURCE_COL_INDEX);
            Cell jsonKeyCell = row.getCell(SHEET3_JSONKEY_COL_INDEX);

            String src = getCellString(srcCell);
            String jsonKey = getCellString(jsonKeyCell);

            if (jsonKey == null || jsonKey.isBlank()) {
                // JSON 키가 없으면 매핑 끝이라고 가정
                break;
            }

            mappings.put(src == null ? "" : src.trim(), jsonKey.trim());
        }
        return mappings;
    }

    private static void clearSheet(Sheet sheet) {
        for (int rowIndex = sheet.getLastRowNum(); rowIndex >= 0; rowIndex--) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    private static void writeSheet4Header(Sheet sheet4) {
        Row headerRow = sheet4.createRow(SHEET4_HEADER_ROW_INDEX);
        headerRow.createCell(SHEET4_COMPANY_IDX_COL_INDEX).setCellValue("COMPANY_IDX");
        headerRow.createCell(SHEET4_JSON_COL_INDEX).setCellValue("JSON");
    }

    private static void saveWorkbookSafely(Workbook workbook, Path excelPath) throws IOException {
        Path parent = excelPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        Path tempPath = excelPath.resolveSibling(excelPath.getFileName() + ".tmp");

        try (OutputStream os = Files.newOutputStream(
                tempPath, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
            workbook.write(os);
        }

        try {
            Files.move(
                    tempPath,
                    excelPath,
                    StandardCopyOption.REPLACE_EXISTING,
                    StandardCopyOption.ATOMIC_MOVE);
        } catch (AtomicMoveNotSupportedException e) {
            try {
                Files.move(tempPath, excelPath, StandardCopyOption.REPLACE_EXISTING);
            } catch (AccessDeniedException ex) {
                throw new IOException("엑셀 파일이 다른 프로세스에서 사용 중입니다. 파일을 닫고 다시 시도하세요: "
                        + excelPath.toAbsolutePath(), ex);
            }
        } catch (AccessDeniedException e) {
            throw new IOException("엑셀 파일이 다른 프로세스에서 사용 중입니다. 파일을 닫고 다시 시도하세요: "
                    + excelPath.toAbsolutePath(), e);
        }
    }

    /**
     * Sheet1의 헤더(2행)를 읽어서 "컬럼명 → 열 인덱스" 맵 생성
     */
    private static Map<String, Integer> buildHeaderIndexMap(Sheet sheet1) {
        Map<String, Integer> map = new HashMap<>();
        Row headerRow = sheet1.getRow(SHEET1_HEADER_ROW_INDEX);
        if (headerRow == null) {
            throw new IllegalStateException("Sheet1의 헤더 행(2행)을 찾을 수 없습니다.");
        }

        short lastCellNum = headerRow.getLastCellNum();
        for (int colIndex = 0; colIndex < lastCellNum; colIndex++) {
            Cell cell = headerRow.getCell(colIndex);
            String name = getCellString(cell);
            if (name != null && !name.isBlank()) {
                map.put(name.trim(), colIndex);
            }
        }
        return map;
    }

    /**
     * 셀을 문자열로 변환 (날짜/숫자/문자 다 문자열로 처리)
     */
    private static String getCellString(Cell cell) {
        if (cell == null) return null;

        CellType type = cell.getCellType();

        switch (type) {
            case STRING:
                return cell.getStringCellValue();

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // 날짜 셀이면 yyyy-MM-dd 형식으로
                    return cell.getLocalDateTimeCellValue()
                            .toLocalDate()
                            .toString();
                } else {
                    double d = cell.getNumericCellValue();
                    if (d == Math.floor(d)) {
                        // 정수면 소수점 없이
                        return String.valueOf((long) d);
                    } else {
                        return String.valueOf(d);
                    }
                }

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());

            case FORMULA:
                // 수식 셀은 결과 타입으로 한 번 평가하고 다시 재귀 호출
                FormulaEvaluator evaluator = cell.getSheet().getWorkbook()
                        .getCreationHelper()
                        .createFormulaEvaluator();
                evaluator.evaluateFormulaCell(cell);
                return getCellString(cell);

            case BLANK:
            case _NONE:
            case ERROR:
            default:
                return "";
        }
    }
}
