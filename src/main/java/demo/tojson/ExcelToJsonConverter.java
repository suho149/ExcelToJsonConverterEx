package demo.tojson;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.AccessDeniedException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ThreadLocalRandom;

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
    private static final int SHEET4_COMPANY_IDX_COL_INDEX = 0;              // A열
    private static final int SHEET4_CONTENTS_COL_INDEX = 1;                 // B열
    private static final int SHEET4_COMPANY_SIZE_COL_INDEX = 2;             // C열
    private static final int SHEET4_DEFENSE_YN_COL_INDEX = 3;               // D열
    private static final int SHEET4_DEFENSE_DATE_COL_INDEX = 4;             // E열
    private static final int SHEET4_DEL_YN_COL_INDEX = 5;                   // F열
    private static final int SHEET4_CREATE_ID_COL_INDEX = 6;                // G열
    private static final int SHEET4_CREATE_IP_COL_INDEX = 7;                // H열
    private static final int SHEET4_CREATE_DATE_COL_INDEX = 8;              // I열
    private static final int SHEET4_UPDATE_ID_COL_INDEX = 9;                // J열
    private static final int SHEET4_UPDATE_IP_COL_INDEX = 10;               // K열
    private static final int SHEET4_UPDATE_DATE_COL_INDEX = 11;             // L열
    private static final int SHEET4_TOTAL_COLUMN_COUNT = 12;
    private static final int EXCEL_MAX_COLUMN_WIDTH = 255 * 256;
    private static final int SHEET4_CONTENTS_FIXED_WIDTH = 60 * 256; // JSON 컬럼 폭 고정

    private static final String DEFAULT_DEFENSE_DESIGNATION_YN = "Y";
    private static final String DEFAULT_DEFENSE_DESIGNATION_DATE = "1900.1.1";
    private static final String DEFAULT_DEL_YN = "N";
    private static final String DEFAULT_CREATE_ID = "root01";
    private static final String DEFAULT_CREATE_IP = "0:0:0:0:0:0:0:756";
    private static final String DEFAULT_UPDATE_VALUE = "\\N";
    private static final DateTimeFormatter CREATE_DATE_FORMATTER = DateTimeFormatter.ofPattern("yy.MM.d");
    private static final DateTimeFormatter SQL_CREATE_DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final DateTimeFormatter SQL_FILE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
    private static final String SQL_OUTPUT_DIR_NAME = "output";
    private static final String SQL_FILE_PREFIX = "sheet4_insert_";
    private static final String SQL_TARGET_TABLE = "T_COMPANY_INFO";
    private static final String RANDOM_COMPANY_IDX_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    private static final int RANDOM_COMPANY_IDX_LENGTH = 6;
    private static final String[] SHEET4_HEADERS = {
            "COMPANY_IDX",
            "CONTENTS",
            "COMPANY_SIZE",
            "DEFENSE_DESIGNATION_YN",
            "DEFENSE_DESIGNATION_DATE",
            "DEL_YN",
            "CREATE_ID",
            "CREATE_IP",
            "CREATE_DATE",
            "UPDATE_ID",
            "UPDATE_IP",
            "UPDATE_DATE"
    };

    private record SqlInsertRow(
            String companyIdx,
            String contents,
            String companySize,
            String defenseDesignationYn,
            String defenseDesignationDate,
            String delYn,
            String createId,
            String createIp,
            String createDate,
            String updateId,
            String updateIp,
            String updateDate) {
    }

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
            List<SqlInsertRow> sqlRows = new ArrayList<>();

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

                String companySizeRaw = getFirstAvailableCellValue(
                        row, headerIndexMap, "company_size", "companySize");
                String companySize = convertCompanySize(companySizeRaw);
                String defenseDesignationYn = getFirstAvailableCellValue(
                        row, headerIndexMap, "defense_designation_yn", "defenseDesignationYn");
                String defenseDesignationDate = getFirstAvailableCellValue(
                        row, headerIndexMap, "defense_designation_date", "defenseDesignationDate");
                String createDate = LocalDate.now().format(CREATE_DATE_FORMATTER);
                String sqlCreateDateTime = LocalDateTime.now().format(SQL_CREATE_DATE_TIME_FORMATTER);
                String resolvedDefenseYn = defaultIfBlank(defenseDesignationYn, DEFAULT_DEFENSE_DESIGNATION_YN);
                String resolvedDefenseDate = defaultIfBlank(defenseDesignationDate, DEFAULT_DEFENSE_DESIGNATION_DATE);

                Row outputRow = sheet4.createRow(outputRowIndex++);
                outputRow.createCell(SHEET4_COMPANY_IDX_COL_INDEX).setCellValue(companyId);
                outputRow.createCell(SHEET4_CONTENTS_COL_INDEX).setCellValue(jsonString);
                outputRow.createCell(SHEET4_COMPANY_SIZE_COL_INDEX).setCellValue(companySize);
                outputRow.createCell(SHEET4_DEFENSE_YN_COL_INDEX).setCellValue(resolvedDefenseYn);
                outputRow.createCell(SHEET4_DEFENSE_DATE_COL_INDEX).setCellValue(resolvedDefenseDate);
                outputRow.createCell(SHEET4_DEL_YN_COL_INDEX).setCellValue(DEFAULT_DEL_YN);
                outputRow.createCell(SHEET4_CREATE_ID_COL_INDEX).setCellValue(DEFAULT_CREATE_ID);
                outputRow.createCell(SHEET4_CREATE_IP_COL_INDEX).setCellValue(DEFAULT_CREATE_IP);
                outputRow.createCell(SHEET4_CREATE_DATE_COL_INDEX).setCellValue(createDate);
                outputRow.createCell(SHEET4_UPDATE_ID_COL_INDEX).setCellValue(DEFAULT_UPDATE_VALUE);
                outputRow.createCell(SHEET4_UPDATE_IP_COL_INDEX).setCellValue(DEFAULT_UPDATE_VALUE);
                outputRow.createCell(SHEET4_UPDATE_DATE_COL_INDEX).setCellValue(DEFAULT_UPDATE_VALUE);

                sqlRows.add(new SqlInsertRow(
                        generateRandomCompanyIdx(),
                        jsonString,
                        companySize,
                        resolvedDefenseYn,
                        null,
                        DEFAULT_DEL_YN,
                        DEFAULT_CREATE_ID,
                        DEFAULT_CREATE_IP,
                        sqlCreateDateTime,
                        null,
                        null,
                        null
                ));
            }

            resizeSheet4Columns(sheet4);

            saveWorkbookSafely(workbook, outputExcelPath);
            Path sqlOutputPath = writeInsertSqlFile(outputExcelPath, sqlRows);

            System.out.println("변환 완료: " + outputExcelPath + " (" + SHEET4_NAME + " 시트)");
            System.out.println("SQL 출력 완료: " + sqlOutputPath);
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
        for (int colIndex = 0; colIndex < SHEET4_HEADERS.length; colIndex++) {
            headerRow.createCell(colIndex).setCellValue(SHEET4_HEADERS[colIndex]);
        }
    }

    private static void resizeSheet4Columns(Sheet sheet4) {
        for (int colIndex = 0; colIndex < SHEET4_TOTAL_COLUMN_COUNT; colIndex++) {
            if (colIndex == SHEET4_CONTENTS_COL_INDEX) {
                continue;
            }

            sheet4.autoSizeColumn(colIndex);

            int minWidthByHeader = Math.min((SHEET4_HEADERS[colIndex].length() + 4) * 256, EXCEL_MAX_COLUMN_WIDTH);
            if (sheet4.getColumnWidth(colIndex) < minWidthByHeader) {
                sheet4.setColumnWidth(colIndex, minWidthByHeader);
            }
        }

        // JSON 컬럼은 autosize 영향 없이 고정 폭으로 유지
        sheet4.setColumnWidth(SHEET4_CONTENTS_COL_INDEX, SHEET4_CONTENTS_FIXED_WIDTH);
    }

    private static Path writeInsertSqlFile(Path excelPath, List<SqlInsertRow> sqlRows) throws IOException {
        Path outputDir = resolveSqlOutputDir(excelPath);
        Files.createDirectories(outputDir);

        String fileName = SQL_FILE_PREFIX + LocalDateTime.now().format(SQL_FILE_TIME_FORMATTER) + ".sql";
        Path sqlPath = outputDir.resolve(fileName);

        try (BufferedWriter writer = Files.newBufferedWriter(
                sqlPath, StandardCharsets.UTF_8, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
            for (SqlInsertRow row : sqlRows) {
                writer.write(buildInsertSql(row));
                writer.newLine();
            }
        }

        return sqlPath.toAbsolutePath().normalize();
    }

    private static Path resolveSqlOutputDir(Path excelPath) {
        Path parent = excelPath.getParent();
        if (parent == null) {
            return Paths.get(SQL_OUTPUT_DIR_NAME);
        }

        Path folderName = parent.getFileName();
        if (folderName != null && "excel".equalsIgnoreCase(folderName.toString())) {
            Path sibling = parent.resolveSibling(SQL_OUTPUT_DIR_NAME);
            if (sibling != null) {
                return sibling;
            }
        }

        return parent.resolve(SQL_OUTPUT_DIR_NAME);
    }

    private static String buildInsertSql(SqlInsertRow row) {
        return "INSERT INTO " + SQL_TARGET_TABLE
                + " (COMPANY_IDX, CONTENTS, COMPANY_SIZE, DEFENSE_DESIGNATION_YN, DEFENSE_DESIGNATION_DATE, "
                + "DEL_YN, CREATE_ID, CREATE_IP, CREATE_DATE, UPDATE_ID, UPDATE_IP, UPDATE_DATE) VALUES ("
                + quoteSql(row.companyIdx()) + ", "
                + quoteSql(row.contents()) + ", "
                + quoteSql(row.companySize()) + ", "
                + quoteSql(row.defenseDesignationYn()) + ", "
                + sqlValueOrNull(row.defenseDesignationDate()) + ", "
                + quoteSql(row.delYn()) + ", "
                + quoteSql(row.createId()) + ", "
                + quoteSql(row.createIp()) + ", "
                + quoteSql(row.createDate()) + ", "
                + sqlValueOrNull(row.updateId()) + ", "
                + sqlValueOrNull(row.updateIp()) + ", "
                + sqlValueOrNull(row.updateDate()) + ");";
    }

    private static String quoteSql(String value) {
        String normalized = value == null ? "" : value;
        return "'" + normalized.replace("'", "''") + "'";
    }

    private static String sqlValueOrNull(String value) {
        if (value == null || value.isBlank()) {
            return "NULL";
        }
        return quoteSql(value);
    }

    private static String generateRandomCompanyIdx() {
        StringBuilder sb = new StringBuilder(RANDOM_COMPANY_IDX_LENGTH);
        for (int i = 0; i < RANDOM_COMPANY_IDX_LENGTH; i++) {
            int randomIndex = ThreadLocalRandom.current().nextInt(RANDOM_COMPANY_IDX_CHARS.length());
            sb.append(RANDOM_COMPANY_IDX_CHARS.charAt(randomIndex));
        }
        return sb.toString();
    }

    private static String getFirstAvailableCellValue(
            Row row, Map<String, Integer> headerIndexMap, String... columnNames) {
        for (String columnName : columnNames) {
            Integer colIndex = headerIndexMap.get(columnName);
            if (colIndex == null) {
                continue;
            }
            String value = getCellString(row.getCell(colIndex));
            if (value != null && !value.isBlank()) {
                return value.trim();
            }
        }
        return "";
    }

    private static String convertCompanySize(String value) {
        if (value == null || value.isBlank()) {
            return "LARGE";
        }
        if ("0".equals(value)) {
            return "LARGE";
        }
        if ("1".equals(value)) {
            return "SMALL";
        }
        return value;
    }

    private static String defaultIfBlank(String value, String defaultValue) {
        if (value == null || value.isBlank()) {
            return defaultValue;
        }
        return value;
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
