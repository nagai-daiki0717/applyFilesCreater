package app;

import com.amazonaws.services.lambda.runtime.Context;
import com.amazonaws.services.lambda.runtime.RequestStreamHandler;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

public class LambdaHandler implements RequestStreamHandler {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private final S3Service s3Service = new S3Service();

    private final TeikiExcelService teikiExcelService = new TeikiExcelService(s3Service);

    private final TeikiExcelReadService teikiExcelReadService = new TeikiExcelReadService();

    private final KotsuhiExcelService kotsuhiExcelService = new KotsuhiExcelService();

    @Override
    public void handleRequest(InputStream input, OutputStream output,
            Context context) {
        try {
            JsonNode event = MAPPER.readTree(input);

            String method = getMethod(event);
            String path = getPath(event);
            String origin = getOrigin(event);

            Map<String, Object> response;

            if ("OPTIONS".equals(method)) {
                response = ApiResponse.json(204, null, origin);
            } else if (Env.BUCKET.isBlank()) {
                response = ApiResponse.json(500, Map.of("error", "config_error",
                        "detail", "TEMP_BUCKET is empty"), origin);
            } else if ("POST".equals(method) && path.endsWith("/upload-url")) {
                response = handleUploadUrl(event, origin);
            } else if ("POST".equals(method) && path.endsWith("/generate")) {
                response = handleGenerate(event, origin);
            } else if ("POST".equals(method) && path.endsWith(
                    "/carfare_generate")) {
                response = handleCarfareGenerate(event, origin);
            } else {
                response = ApiResponse.json(404, Map.of("error", "not_found",
                        "path", path, "method", method), origin);
            }

            MAPPER.writeValue(output, response);

        } catch (Exception e) {
            try {
                MAPPER.writeValue(output, ApiResponse.json(500, Map.of("error",
                        "internal_error", "detail", e.toString()), "*"));
            } catch (Exception ignored) {
            }
        }
    }

    private Map<String, Object> handleUploadUrl(JsonNode event, String origin) {
        JsonNode payload = readJsonBody(event);

        String fileName = payload.path("file_name").asText("upload.bin");
        String contentType = payload.path("content_type").asText(
                "application/octet-stream");

        String ext = "";
        int idx = fileName.lastIndexOf('.');
        if (idx >= 0) {
            ext = fileName.substring(idx).toLowerCase();
        }

        String yyyymmdd = LocalDate.now(ZoneOffset.UTC).format(
                DateTimeFormatter.BASIC_ISO_DATE);
        String key = Env.TEMP_PREFIX + "/" + yyyymmdd + "/" + UUID.randomUUID()
                .toString().replace("-", "") + ext;

        String putUrl = s3Service.presignPutUrl(Env.BUCKET, key, contentType,
                600);

        return ApiResponse.json(200, Map.of("put_url", putUrl, "file_key", key,
                "expires_in", 600), origin);
    }

    private Map<String, Object> handleGenerate(JsonNode event, String origin) {
        Path templateLocal = Path.of("/tmp/template.xlsx");
        Path previousExcelLocal = Path.of("/tmp/previous_application.xlsx");
        Path outputLocal = Path.of("/tmp/output.xlsx");

        Set<String> tempUploadKeys = new LinkedHashSet<>();

        try {
            if (Env.BUCKET.isBlank() || Env.TEMPLATE_KEY.isBlank()) {
                return ApiResponse.json(500, Map.of("error", "config_error",
                        "detail", "TEMPLATE_BUCKET or TEMPLATE_KEY is empty"),
                        origin);
            }
            if (Env.BUCKET.isBlank()) {
                return ApiResponse.json(500, Map.of("error", "config_error",
                        "detail", "BUCKET is empty"), origin);
            }

            JsonNode payload = readJsonBody(event);
            tempUploadKeys.addAll(collectTemporaryUploadKeys(payload));

            String previousExcelKey = payload.path("commute").path(
                    "previous_application_excel").path("file_key").asText("");

            if (!previousExcelKey.isBlank()) {
                if (!isTemporaryUploadKey(previousExcelKey)) {
                    return ApiResponse.json(400, Map.of("error",
                            "invalid_file_key", "detail",
                            "previous_application_excel.file_key is not under temporary prefix"),
                            origin);
                }

                s3Service.downloadToFile(Env.BUCKET, previousExcelKey,
                        previousExcelLocal);
                payload = teikiExcelReadService.buildPayloadFromPreviousExcel(
                        previousExcelLocal, payload);
            }

            s3Service.downloadToFile(Env.BUCKET, Env.TEMPLATE_KEY,
                    templateLocal);

            java.util.List<String> usedPhotoKeys = teikiExcelService
                    .generateWorkbook(templateLocal, outputLocal, payload);
            tempUploadKeys.addAll(usedPhotoKeys);

            String yyyymmdd = LocalDate.now(ZoneOffset.UTC).format(
                    DateTimeFormatter.BASIC_ISO_DATE);
            String outKey = Env.OUTPUT_PREFIX + "/" + yyyymmdd + "/" + UUID
                    .randomUUID().toString().replace("-", "") + ".xlsx";

            s3Service.uploadFile(outputLocal, Env.BUCKET, outKey,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            String downloadFilename = teikiExcelService.buildOutputFilename(
                    payload);
            String downloadUrl = s3Service.presignGetUrl(Env.BUCKET, outKey,
                    teikiExcelService.contentDisposition(downloadFilename),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    600);

            return ApiResponse.json(200, Map.of("download_url", downloadUrl,
                    "file_key", outKey, "expires_in", 600), origin);

        } catch (Exception e) {
            return ApiResponse.json(500, Map.of("error", "internal_error",
                    "detail", e.toString()), origin);

        } finally {
            deleteTemporaryUploadObjects(tempUploadKeys);

            try {
                java.nio.file.Files.deleteIfExists(templateLocal);
            } catch (Exception ex) {
                System.out.println("failed to delete local template: " + ex);
            }

            try {
                java.nio.file.Files.deleteIfExists(previousExcelLocal);
            } catch (Exception ex) {
                System.out.println("failed to delete local previous excel: "
                        + ex);
            }

            try {
                java.nio.file.Files.deleteIfExists(outputLocal);
            } catch (Exception ex) {
                System.out.println("failed to delete local output: " + ex);
            }
        }
    }

    private Map<String, Object> handleCarfareGenerate(JsonNode event,
            String origin) {
        Path templateLocal = Path.of("/tmp/carfare_template.xlsx");
        Path outputLocal = Path.of("/tmp/output.xlsx");

        try {
            if (Env.BUCKET.isBlank() || Env.CARFARE_TEMPLATE_KEY.isBlank()) {
                return ApiResponse.json(500, Map.of("error", "config_error",
                        "detail", "BUCKET or CARFARE_TEMPLATE_KEY is empty"),
                        origin);
            }

            JsonNode payload = readJsonBody(event);

            System.out.println("carfare generate start");
            System.out.println("template bucket=" + Env.BUCKET + ", key="
                    + Env.CARFARE_TEMPLATE_KEY);

            s3Service.downloadToFile(Env.BUCKET, Env.CARFARE_TEMPLATE_KEY,
                    templateLocal);

            kotsuhiExcelService.generateWorkbook(templateLocal, outputLocal,
                    payload);

            String yyyymmdd = LocalDate.now(ZoneOffset.UTC).format(
                    DateTimeFormatter.BASIC_ISO_DATE);

            String outKey = Env.OUTPUT_PREFIX + "/" + yyyymmdd + "/" + UUID
                    .randomUUID().toString().replace("-", "") + ".xlsx";

            System.out.println("upload output bucket=" + Env.BUCKET + ", key="
                    + outKey);

            s3Service.uploadFile(outputLocal, Env.BUCKET, outKey,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            String downloadFilename = kotsuhiExcelService.buildOutputFilename(
                    payload);

            String downloadUrl = s3Service.presignGetUrl(Env.BUCKET, outKey,
                    kotsuhiExcelService.contentDisposition(downloadFilename),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    600);

            return ApiResponse.json(200, Map.of("download_url", downloadUrl,
                    "file_key", outKey, "expires_in", 600), origin);

        } catch (Exception e) {
            e.printStackTrace();

            return ApiResponse.json(500, Map.of("error", "internal_error",
                    "detail", e.toString()), origin);

        } finally {
            try {
                java.nio.file.Files.deleteIfExists(templateLocal);
            } catch (Exception ex) {
                System.out.println("failed to delete local template: " + ex);
            }

            try {
                java.nio.file.Files.deleteIfExists(outputLocal);
            } catch (Exception ex) {
                System.out.println("failed to delete local output: " + ex);
            }
        }
    }

    private Set<String> collectTemporaryUploadKeys(JsonNode payload) {
        Set<String> keys = new LinkedHashSet<>();

        JsonNode previousExcel = payload.path("commute").path(
                "previous_application_excel");
        addIfTemporaryUploadKey(keys, previousExcel.path("file_key").asText(
                ""));

        JsonNode passPhotos = payload.path("commute").path("pass_photos");
        if (passPhotos.isArray()) {
            for (JsonNode photo : passPhotos) {
                addIfTemporaryUploadKey(keys, photo.path("file_key").asText(
                        ""));
            }
        }

        return keys;
    }

    private void addIfTemporaryUploadKey(Set<String> keys, String key) {
        if (isTemporaryUploadKey(key)) {
            keys.add(key);
        }
    }

    private boolean isTemporaryUploadKey(String key) {
        if (key == null || key.isBlank() || Env.TEMP_PREFIX.isBlank()) {
            return false;
        }
        return key.startsWith(Env.TEMP_PREFIX + "/");
    }

    private void deleteTemporaryUploadObjects(Set<String> keys) {
        for (String key : keys) {
            if (!isTemporaryUploadKey(key)) {
                System.out.println("skip delete non-temp upload key: " + key);
                continue;
            }

            try {
                s3Service.deleteObject(Env.BUCKET, key);
                System.out.println("deleted temp upload object: " + key);
            } catch (Exception ex) {
                System.out.println("failed to delete temp upload object: " + key
                        + " / " + ex);
            }
        }
    }

    private JsonNode readJsonBody(JsonNode event) {
        try {
            String raw = event.path("body").asText("{}");
            boolean isBase64 = event.path("isBase64Encoded").asBoolean(false);
            if (isBase64) {
                raw = new String(java.util.Base64.getDecoder().decode(
                        raw), StandardCharsets.UTF_8);
            }
            return MAPPER.readTree(raw);
        } catch (Exception e) {
            return MAPPER.createObjectNode();
        }
    }

    private String getMethod(JsonNode event) {
        String method = event.path("requestContext").path("http").path("method")
                .asText("");
        if (method.isBlank()) {
            method = event.path("httpMethod").asText("POST");
        }
        return method.toUpperCase();
    }

    private String getPath(JsonNode event) {
        String path = event.path("requestContext").path("http").path("path")
                .asText("");
        if (path.isBlank()) {
            path = event.path("path").asText("");
        }
        return path.endsWith("/") && path.length() > 1 ? path.substring(0, path
                .length() - 1) : path;
    }

    private String getOrigin(JsonNode event) {
        JsonNode headers = event.path("headers");
        String origin = headers.path("origin").asText("");
        if (origin.isBlank()) {
            origin = headers.path("Origin").asText("*");
        }
        return origin;
    }
}
