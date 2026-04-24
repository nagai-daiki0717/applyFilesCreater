package app;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import org.apache.commons.lang3.StringUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;

public class TeikiExcelService {

    private final S3Service s3Service;

    public TeikiExcelService(S3Service s3Service) {
        this.s3Service = s3Service;
    }

    public java.util.List<String> generateWorkbook(Path templatePath,
            Path outputPath, JsonNode payload) throws Exception {
        FileInputStream fis = null;
        XSSFWorkbook workbook = null;
        FileOutputStream fos = null;

        try {
            fis = new FileInputStream(templatePath.toFile());
            workbook = new XSSFWorkbook(fis);

            java.util.List<String> tempImageKeys = applyCommuteApplicationToWorkbook(
                    workbook, payload);

            FormulaEvaluator evaluator = workbook.getCreationHelper()
                    .createFormulaEvaluator();
            evaluator.evaluateAll();

            fos = new FileOutputStream(outputPath.toFile());
            workbook.write(fos);

            return tempImageKeys;
        } finally {
            if (fos != null)
                fos.close();
            if (workbook != null)
                workbook.close();
            if (fis != null)
                fis.close();
        }
    }

    private java.util.List<String> applyCommuteApplicationToWorkbook(
            XSSFWorkbook wb, JsonNode payload) throws Exception {
        LocalDate submitDate = parseDate(text(payload, "submitted_at"));
        LocalDate reasonDate = parseDate(text(payload.path("notification"),
                "reason_effective_date"));

        String workplaceName = text(payload.path("work_location"), "name");
        String workplaceAddr = text(payload.path("work_location"), "address");
        String applicantName = text(payload.path("applicant"), "name");
        String reasonText = text(payload.path("notification"), "reason_text");

        JsonNode commute = payload.path("commute");
        JsonNode passes = commute.path("passes");
        JsonNode passPhotos = commute.path("pass_photos");

        int totalMinutes = 0;
        if (passes.isArray()) {
            for (JsonNode p : passes) {
                totalMinutes += p.path("one_way_minutes").asInt(0);
            }
        }

        XSSFSheet wsMain = wb.getSheet("通勤交通費申請書");
        if (wsMain == null) {
            throw new IllegalStateException("sheet not found: 通勤交通費申請書");
        }

        if (submitDate != null) {
            setCell(wsMain, "K3", formatSubmitDateReiwa(submitDate));
        }
        setCell(wsMain, "C5", workplaceName);
        setCell(wsMain, "C6", workplaceAddr);
        setCell(wsMain, "C7", applicantName);

        String originalM5 = getCellString(wsMain, "M5");
        setCell(wsMain, "M5", rewriteReasonBlock(originalM5, reasonText,
                reasonDate));

        int startRow = 10;
        int endRow = 15;
        int maxRows = endRow - startRow + 1;

        for (int i = 0; i < maxRows; i++) {
            int r = startRow + i;
            JsonNode p = (passes.isArray() && i < passes.size()) ? passes.get(i)
                    : null;

            String transportation = p == null ? ""
                    : p.path("transportation").asText("");
            String section = p == null ? "" : p.path("section").asText("");
            int minutes = p == null ? 0 : p.path("one_way_minutes").asInt(0);
            JsonNode amountNode = p == null ? null : p.get("amount_yen");
            String note = p == null ? "" : p.path("note").asText("");

            setCell(wsMain, "B" + r, transportation);
            setCell(wsMain, "D" + r, section);
            setCell(wsMain, "H" + r, minutes == 0 ? "" : minutes + "分");
            setCell(wsMain, "I" + r, "1ヶ月");
            setCellPreserveType(wsMain, Env.PASSES_PRICE_COL + r, amountNode);
            setCell(wsMain, "N" + r, note);
        }

        setCell(wsMain, "E16", minutesToHM(totalMinutes));

        XSSFSheet wsHist = wb.getSheet("定期購入履歴");
        if (wsHist == null) {
            throw new IllegalStateException("sheet not found: 定期購入履歴");
        }

        if (submitDate != null) {
            setCell(wsHist, "D6", submitDate.format(DateTimeFormatter.ofPattern(
                    "yyyy/MM/dd")));
        }

        return addPassPhotos(wb, wsHist, passPhotos);
    }

    private java.util.List<String> addPassPhotos(XSSFWorkbook wb, XSSFSheet ws,
            JsonNode passPhotos) throws Exception {
        java.util.List<String> usedKeys = new java.util.ArrayList<>();

        if (passPhotos == null || !passPhotos.isArray() || passPhotos
                .isEmpty()) {
            return usedKeys;
        }

        int photoCount = Math.min(passPhotos.size(), 6);

        int rowStep;
        int targetWidthPx;
        int rightCol; // 0始まりでI列

        // 同一申請内では全画像共通の幅
        if (photoCount <= 2) {
            rightCol = 8;
            rowStep = 0;
            targetWidthPx = 160;
        } else if (photoCount <= 4) {
            rightCol = 5;
            rowStep = 5;
            targetWidthPx = 120;
        } else {
            rightCol = 4;
            rowStep = 4;
            targetWidthPx = 92;
        }

        int baseRow = 6; // 0始まりで7行目
        int leftCol = 1; // 0始まりでB列

        Drawing<?> drawing = ws.createDrawingPatriarch();

        for (int idx = 0; idx < photoCount; idx++) {
            JsonNode info = passPhotos.get(idx);
            String key = info.path("file_key").asText("");
            if (key.isBlank()) {
                continue;
            }

            byte[] imageBytes = s3Service.downloadBytes(Env.BUCKET, key);
            int format = detectPictureType(key);
            int pictureIndex = wb.addPicture(imageBytes, format);

            ImageDimension dim = readImageDimension(imageBytes);

            // 元画像の縦横比を維持
            double scale = (double) targetWidthPx / dim.width;
            int drawWidthPx = targetWidthPx;
            int drawHeightPx = Math.max(1, (int) Math.round(dim.height
                    * scale));

            int block = idx / 2;
            boolean isLeft = (idx % 2 == 0);

            int startRow = baseRow + block * rowStep;
            int startCol = isLeft ? leftCol : rightCol;

            AnchorPos endCol = calcAnchorEndCol(ws, startCol, drawWidthPx);
            AnchorPos endRow = calcAnchorEndRow(ws, startRow, drawHeightPx);

            ClientAnchor anchor = wb.getCreationHelper().createClientAnchor();
            anchor.setCol1(startCol);
            anchor.setRow1(startRow);
            anchor.setDx1(0);
            anchor.setDy1(0);

            anchor.setCol2(endCol.index);
            anchor.setDx2(endCol.emuOffset);

            anchor.setRow2(endRow.index);
            anchor.setDy2(endRow.emuOffset);

            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

            drawing.createPicture(anchor, pictureIndex);

            usedKeys.add(key);
        }

        return usedKeys;
    }

    private ImageDimension readImageDimension(
            byte[] imageBytes) throws Exception {
        try (InputStream is = new java.io.ByteArrayInputStream(imageBytes)) {
            javax.imageio.stream.ImageInputStream iis = javax.imageio.ImageIO
                    .createImageInputStream(is);
            Iterator<javax.imageio.ImageReader> readers = javax.imageio.ImageIO
                    .getImageReaders(iis);
            if (!readers.hasNext()) {
                return new ImageDimension(100, 100);
            }
            javax.imageio.ImageReader reader = readers.next();
            try {
                reader.setInput(iis);
                return new ImageDimension(reader.getWidth(0), reader.getHeight(
                        0));
            } finally {
                reader.dispose();
                iis.close();
            }
        }
    }

    private static class AnchorPos {
        final int index;

        final int emuOffset;

        AnchorPos(int index, int emuOffset) {
            this.index = index;
            this.emuOffset = emuOffset;
        }
    }

    private int detectPictureType(String key) {
        String lower = key.toLowerCase();
        if (lower.endsWith(".png"))
            return Workbook.PICTURE_TYPE_PNG;
        if (lower.endsWith(".jpg") || lower.endsWith(".jpeg"))
            return Workbook.PICTURE_TYPE_JPEG;
        if (lower.endsWith(".dib"))
            return Workbook.PICTURE_TYPE_DIB;
        if (lower.endsWith(".emf"))
            return Workbook.PICTURE_TYPE_EMF;
        if (lower.endsWith(".wmf"))
            return Workbook.PICTURE_TYPE_WMF;
        return Workbook.PICTURE_TYPE_PNG;
    }

    private void setCell(XSSFSheet sheet, String ref, String value) {
        var cellRef = new org.apache.poi.ss.util.CellReference(ref);
        Row row = sheet.getRow(cellRef.getRow());
        if (row == null)
            row = sheet.createRow(cellRef.getRow());
        var cell = row.getCell(cellRef.getCol());
        if (cell == null)
            cell = row.createCell(cellRef.getCol());
        cell.setCellValue(value == null ? "" : value);
    }

    private void setCellPreserveType(XSSFSheet sheet, String ref,
            JsonNode valueNode) {
        var cellRef = new org.apache.poi.ss.util.CellReference(ref);
        Row row = sheet.getRow(cellRef.getRow());
        if (row == null)
            row = sheet.createRow(cellRef.getRow());
        var cell = row.getCell(cellRef.getCol());
        if (cell == null)
            cell = row.createCell(cellRef.getCol());

        if (valueNode == null || valueNode.isNull()) {
            cell.setBlank();
            return;
        }
        if (valueNode.isNumber()) {
            cell.setCellValue(valueNode.asDouble());
        } else {
            cell.setCellValue(valueNode.asText(""));
        }
    }

    private String getCellString(XSSFSheet sheet, String ref) {
        var cellRef = new org.apache.poi.ss.util.CellReference(ref);
        Row row = sheet.getRow(cellRef.getRow());
        if (row == null)
            return "";
        var cell = row.getCell(cellRef.getCol());
        if (cell == null)
            return "";
        return cell.toString();
    }

    private String text(JsonNode node, String field) {
        JsonNode child = node == null ? null : node.get(field);
        return child == null || child.isNull() ? "" : child.asText("");
    }

    private LocalDate parseDate(String s) {
        if (s == null || s.isBlank())
            return null;
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

    private String rewriteReasonBlock(String originalText, String reasonText,
            LocalDate reasonDate) {
        String[] lines = originalText == null ? new String[0]
                : originalText.split("\\R", -1);

        if (reasonText != null && !reasonText.isBlank()) {
            for (int i = 0; i < lines.length; i++) {
                if (lines[i].contains(reasonText)) {
                    int idx = lines[i].indexOf('□');
                    if (idx >= 0) {
                        lines[i] = lines[i].substring(0, idx) + "☑" + lines[i]
                                .substring(idx + 1);
                    }
                    break;
                }
            }
        }

        if (reasonDate != null) {
            String header = "上記理由が生じた年月日";
            boolean replaced = false;

            for (int i = 0; i < lines.length; i++) {
                if (lines[i].contains(header)) {
                    String insert = String.format("　  %d年　%d月　%d日",
                            reasonDate.getYear(), reasonDate.getMonthValue(),
                            reasonDate.getDayOfMonth());
                    if (i + 1 < lines.length) {
                        lines[i + 1] = insert;
                    }
                    replaced = true;
                    break;
                }
            }

            if (!replaced) {
                String insert = String.format("上記理由が生じた年月日%n　　%d年　%d月　%d日",
                        reasonDate.getYear(), StringUtils.leftPad(String.valueOf(reasonDate.getMonthValue()), 3, " "),
                        reasonDate.getDayOfMonth());
                return (originalText == null || originalText.isBlank()) ? insert
                        : originalText + System.lineSeparator() + insert;
            }
        }

        return String.join("\n", lines);
    }

    private Integer toReiwa(LocalDate date) {
        LocalDate reiwaStart = LocalDate.of(2019, 5, 1);
        if (date.isBefore(reiwaStart))
            return null;
        return date.getYear() - 2018;
    }

    private String formatSubmitDateReiwa(LocalDate submitDate) {
        Integer reiwaYear = toReiwa(submitDate);
        if (reiwaYear == null) {
            return String.format("提出日　　%d年　%d月　%d日", submitDate.getYear(),
                    submitDate.getMonthValue(), submitDate.getDayOfMonth());
        }
        return String.format("提出日　　令和%d年　%d月　%d日", reiwaYear, submitDate
                .getMonthValue(), submitDate.getDayOfMonth());
    }

    private String minutesToHM(int minutes) {
        int h = minutes / 60;
        int m = minutes % 60;
        return h + "時間　　　" + m + "分";
    }

    public String buildOutputFilename(JsonNode payload) {
        LocalDate submitDate = parseDate(text(payload, "submitted_at"));
        String yyyymm = submitDate == null ? "日付不明"
                : String.format("%d年%02d月", submitDate.getYear(), submitDate
                        .getMonthValue());

        String applicant = normalizeApplicantName(text(payload.path(
                "applicant"), "name"));
        return "定期申請書" + yyyymm + "(" + applicant + ").xlsx";
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
        if (name == null || name.isBlank())
            return "名無し";
        String normalized = name.replaceAll("\\s+", "").replace("\u3000", "");
        return normalized.isBlank() ? "名無し" : normalized;
    }

    private static class ImageDimension {
        final int width;

        final int height;

        ImageDimension(int width, int height) {
            this.width = Math.max(1, width);
            this.height = Math.max(1, height);
        }
    }

    private int getColumnWidthPx(XSSFSheet ws, int colIndex0Based) {
        int poiWidth = ws.getColumnWidth(colIndex0Based);
        return (int) Math.floor((poiWidth / 256.0) * 7 + 5);
    }

    private int getRowHeightPx(XSSFSheet ws, int rowIndex0Based) {
        Row row = ws.getRow(rowIndex0Based);
        float heightPt = row != null ? row.getHeightInPoints()
                : ws.getDefaultRowHeightInPoints();
        return Math.round(heightPt * 96 / 72f);
    }

    private AnchorPos calcAnchorEndCol(XSSFSheet ws, int startCol0Based,
            int widthPx) {
        int remain = widthPx;
        int col = startCol0Based;

        while (true) {
            int colWidthPx = getColumnWidthPx(ws, col);
            if (remain <= colWidthPx) {
                return new AnchorPos(col, remain * 9525);
            }
            remain -= colWidthPx;
            col++;
        }
    }

    private AnchorPos calcAnchorEndRow(XSSFSheet ws, int startRow0Based,
            int heightPx) {
        int remain = heightPx;
        int row = startRow0Based;

        while (true) {
            int rowHeightPx = getRowHeightPx(ws, row);
            if (remain <= rowHeightPx) {
                return new AnchorPos(row, remain * 9525);
            }
            remain -= rowHeightPx;
            row++;
        }
    }
}
