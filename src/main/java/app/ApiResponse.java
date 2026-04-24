package app;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.util.LinkedHashMap;
import java.util.Map;

public final class ApiResponse {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private ApiResponse() {
    }

    public static Map<String, Object> json(int statusCode, Object body,
            String origin) {
        Map<String, Object> res = new LinkedHashMap<>();
        res.put("statusCode", statusCode);

        Map<String, String> headers = new LinkedHashMap<>();
        headers.put("Content-Type", "application/json");
        headers.put("Access-Control-Allow-Origin", origin == null || origin
                .isBlank() ? "*" : origin);
        headers.put("Access-Control-Allow-Headers", "Content-Type");
        headers.put("Access-Control-Allow-Methods", "POST,OPTIONS");
        headers.put("Access-Control-Max-Age", "600");
        res.put("headers", headers);

        try {
            res.put("body", body == null ? ""
                    : MAPPER.writeValueAsString(body));
        } catch (Exception e) {
            res.put("body", "{\"error\":\"serialize_failed\"}");
        }
        return res;
    }
}
