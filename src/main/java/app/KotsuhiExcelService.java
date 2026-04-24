package app;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class KotsuhiExcelService {

    public java.util.List<String> generateWorkbook(Path templatePath,
            Path outputPath, JsonNode payload) throws Exception {
        FileInputStream fis = null;
        XSSFWorkbook workbook = null;
        FileOutputStream fos = null;

        try {
            fis = new FileInputStream(templatePath.toFile());
            workbook = new XSSFWorkbook(fis);

            java.util.List<String> tempImageKeys = applyCommuteApplicationToWorkbook(workbook, payload);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fos = new FileOutputStream(outputPath.toFile());
            workbook.write(fos);

            return tempImageKeys;
        } finally {
            if (fos != null) fos.close();
            if (workbook != null) workbook.close();
            if (fis != null) fis.close();
        }
    }

    private java.util.List<String> applyCommuteApplicationToWorkbook(
            XSSFWorkbook wb, JsonNode payload) throws Exception {

        LocalDate submitDate = parseDate(text(payload, "submit_date"));

        String applicantNo = text(payload, "applicant_no");
        String applicantName = text(payload, "applicant_name");
        String applicantProjectName = text(payload, "applicant_project_name");

        JsonNode carfareDetails = payload.path("carfare_details");

        XSSFSheet wsMain = wb.getSheet("交通費精算書");
        if (wsMain == null) {
            wsMain = wb.getSheetAt(0);
        }

        // ===== ヘッダ部 =====
        // テンプレートのセル位置に応じて調整してください
        if (submitDate != null) {
            setCell(wsMain, "AD2", submitDate.format(DateTimeFormatter.ofPattern("yyyy年MM月dd日")));
        }
        setCell(wsMain, "E10", submitDate.format(DateTimeFormatter.ofPattern("M")));
        setCell(wsMain, "B5", applicantNo);
        setCell(wsMain, "H5", applicantName);
        setCell(wsMain, "H7", applicantProjectName);

        // ===== 明細部 =====
        int startRow = 13;

        for (int i = 0; i < carfareDetails.size(); i++) {
            int rowNum = startRow + i;
            JsonNode detail = carfareDetails.get(i);

            String carfareDate = detail.path("carfare_date").asText("");
            String destination = detail.path("carfare_destination").asText("");
            String line = detail.path("carfare_destination_line").asText("");
            String reason = detail.path("carfare_reason").asText("");
            String departure = detail.path("carfare_departure").asText("");
            String arrival = detail.path("carfare_arrival").asText("");
            int unitPrice = detail.path("carfare_unit_price").asInt();
            String tripType = detail.path("carfare_one_way_or_round_trip").asText("");
            String section = detail.path("carfare_section").asText("");

            // テンプレートの列位置に応じて調整してください
            setCell(wsMain, "C" + rowNum, formatMd(carfareDate));      // 月日
            setCell(wsMain, "E" + rowNum, destination);               // 行き先
            setCell(wsMain, "I" + rowNum, line);                      // 利用路線（機関）
            setCell(wsMain, "M" + rowNum, reason);                    // 理由
            setCell(wsMain, "S" + rowNum, departure);                 // 出発
            setCell(wsMain, "V" + rowNum, arrival);                   // 到着
            setCell(wsMain, "Y" + rowNum, String.valueOf(unitPrice)); // 単価
            setCell(wsMain, "AA" + rowNum, toTripLabel(tripType));     // 片/往
            setCell(wsMain, "AC" + rowNum, toSectionLabel(section));   // 交通費区分
        }

        // 合計金額はテンプレートの数式で算出するため、setCellしない
        return new java.util.ArrayList<>();
    }

    private String toTripLabel(String tripType) {
        if ("round_trip".equals(tripType)) {
            return "往";
        }
        return "片";
    }

    private String toSectionLabel(String section) {
        switch (section) {
            case "return_to_office":
                return "帰社";
            case "remote_within_company":
                return "リモート【社内】";
            case "remote_field_work":
                return "リモート【現場】";
            case "other":
                return "その他";
            case "customer_site":
                return "顧客先";
            default:
                return section;
        }
    }

    private String formatMd(String dateText) {
        LocalDate d = parseDate(dateText);
        if (d == null) {
            return dateText == null ? "" : dateText;
        }
        return d.format(DateTimeFormatter.ofPattern("M/d"));
    }

    private void setCell(XSSFSheet sheet, String ref, String value) {
        var cellRef = new org.apache.poi.ss.util.CellReference(ref);
        Row row = sheet.getRow(cellRef.getRow());
        if (row == null) row = sheet.createRow(cellRef.getRow());
        var cell = row.getCell(cellRef.getCol());
        if (cell == null) cell = row.createCell(cellRef.getCol());
        cell.setCellValue(value == null ? "" : value);
    }

    private String text(JsonNode node, String field) {
        JsonNode child = node == null ? null : node.get(field);
        return child == null || child.isNull() ? "" : child.asText("");
    }

    private LocalDate parseDate(String s) {
        if (s == null || s.isBlank()) return null;
        s = s.trim();
        try {
            if (s.length() == 7 && s.charAt(4) == '-') {
                return LocalDate.parse(s + "-01");
            }
            return LocalDate.parse(s.substring(0, 10));
        } catch (Exception e) {
            return null;
        }
    }

    public String buildOutputFilename(JsonNode payload) {
        LocalDate submitDate = parseDate(text(payload, "submit_date"));
        String yyyymm = String.format("%d年度%02d月", submitDate.getYear(), submitDate.getMonthValue());

        String applicant = normalizeApplicantName(text(payload, "applicant_name"));
        return "交通費精算書" + yyyymm + "(" + applicant + ").xlsx";
    }

    public String contentDisposition(String filename) {
        String asciiFallback = "download.xlsx";
        String encoded = java.net.URLEncoder.encode(
                filename,
                java.nio.charset.StandardCharsets.UTF_8
        ).replace("+", "%20");

        return "attachment; filename=\"" + asciiFallback + "\"; filename*=UTF-8''" + encoded;
    }

    private String normalizeApplicantName(String name) {
        if (name == null || name.isBlank()) return "名無し";
        String normalized = name.replaceAll("\\s+", "").replace("\u3000", "");
        return normalized.isBlank() ? "名無し" : normalized;
    }
}