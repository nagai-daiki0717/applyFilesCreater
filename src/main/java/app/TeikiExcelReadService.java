package app;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.nio.file.Path;

public class TeikiExcelReadService {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    public ObjectNode buildPayloadFromPreviousExcel(Path previousExcelPath,
            JsonNode requestPayload) throws Exception {
        try (FileInputStream fis = new FileInputStream(previousExcelPath.toFile());
                XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet wsMain = wb.getSheet("通勤交通費申請書");
            if (wsMain == null) {
                throw new IllegalStateException("sheet not found: 通勤交通費申請書");
            }

            ObjectNode merged = MAPPER.createObjectNode();
            copyText(merged, requestPayload, "request_id");
            merged.put("input_mode", "previous_excel");
            merged.put("submitted_at", text(requestPayload, "submitted_at"));

            ObjectNode workLocation = merged.putObject("work_location");
            workLocation.put("name", cellString(wsMain, "C5"));
            workLocation.put("address", cellString(wsMain, "C6"));

            ObjectNode applicant = merged.putObject("applicant");
            applicant.put("name", cellString(wsMain, "C7"));

            ObjectNode notification = merged.putObject("notification");
            notification.put("reason_text", extractCheckedReason(cellString(wsMain, "M5")));
            notification.put("reason_effective_date", text(requestPayload.path("notification"),
                    "reason_effective_date"));

            ObjectNode commute = merged.putObject("commute");
            commute.put("usage_start_date", text(requestPayload.path("commute"),
                    "usage_start_date"));
            commute.set("passes", readPasses(wsMain));
            commute.set("pass_photos", requestPayload.path("commute").path("pass_photos")
                    .deepCopy());
            commute.set("previous_application_excel", requestPayload.path("commute")
                    .path("previous_application_excel").deepCopy());

            return merged;
        }
    }

    private ArrayNode readPasses(XSSFSheet wsMain) {
        ArrayNode passes = MAPPER.createArrayNode();

        int startRow = 10;
        int endRow = 15;

        for (int r = startRow; r <= endRow; r++) {
            String transportation = cellString(wsMain, "B" + r);
            String section = cellString(wsMain, "D" + r);
            String minutesText = cellString(wsMain, "H" + r);
            String purchasePeriod = cellString(wsMain, "I" + r);
            String amountText = cellString(wsMain, Env.PASSES_PRICE_COL + r);
            String note = cellString(wsMain, "N" + r);

            if (transportation.isBlank() && section.isBlank() && amountText.isBlank()) {
                continue;
            }

            ObjectNode pass = MAPPER.createObjectNode();
            pass.put("transportation", transportation);
            pass.put("section", section);
            pass.put("one_way_minutes", extractInt(minutesText));
            pass.put("purchase_period_text", purchasePeriod.isBlank() ? "1ヶ月" : purchasePeriod);
            pass.put("amount_yen", extractInt(amountText));
            pass.put("note", note);
            passes.add(pass);
        }

        return passes;
    }

    private String extractCheckedReason(String reasonBlock) {
        if (reasonBlock == null || reasonBlock.isBlank()) {
            return "";
        }

        String[] reasonCandidates = {
                "新規",
                "異動",
                "転居",
                "通勤経路・方法の変更",
                "運賃の負担額の変更",
                "継続"
        };

        for (String line : reasonBlock.split("\\R")) {
            if (line.contains("☑") || line.contains("■") || line.contains("✓") || line.contains("✔")) {
                for (String reason : reasonCandidates) {
                    if (line.contains(reason)) {
                        return reason;
                    }
                }
            }
        }
        return "";
    }

    private int extractInt(String value) {
        if (value == null || value.isBlank()) {
            return 0;
        }
        String digits = value.replaceAll("[^0-9-]", "");
        if (digits.isBlank() || "-".equals(digits)) {
            return 0;
        }
        try {
            return Integer.parseInt(digits);
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private String cellString(XSSFSheet sheet, String ref) {
        var cellRef = new org.apache.poi.ss.util.CellReference(ref);
        Row row = sheet.getRow(cellRef.getRow());
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(cellRef.getCol());
        if (cell == null) {
            return "";
        }
        return new DataFormatter().formatCellValue(cell).trim();
    }

    private void copyText(ObjectNode target, JsonNode source, String field) {
        String value = text(source, field);
        if (!value.isBlank()) {
            target.put(field, value);
        }
    }

    private String text(JsonNode node, String field) {
        JsonNode child = node == null ? null : node.get(field);
        return child == null || child.isNull() ? "" : child.asText("");
    }
}
