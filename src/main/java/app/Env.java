package app;

public final class Env {

    public static final String REGION = getenv("AWS_REGION", "ap-northeast-1");

    public static final String TEMP_PREFIX = trimSlash(getenv("TEMP_PREFIX",
            "tmp"));

    public static final String BUCKET = getenv("BUCKET", "");

    public static final String TEMPLATE_KEY = getenv("TEMPLATE_KEY",
            "template/定期申請書YYYY年度MM月(氏名).xlsx");
    
    public static final String CARFARE_TEMPLATE_KEY = getenv("CARFARE_TEMPLATE_KEY",
            "template/交通費精算書YYYY年度MM月(氏名).xlsx");
    
    public static final String OUTPUT_PREFIX = trimSlash(getenv("OUTPUT_PREFIX",
            "temp"));

    public static final String PASSES_PRICE_COL = getenv("PASSES_PRICE_COL",
            "K").toUpperCase();

    private Env() {
    }

    private static String getenv(String key, String defaultValue) {
        String value = System.getenv(key);
        return value == null || value.isBlank() ? defaultValue : value;
    }

    private static String trimSlash(String s) {
        if (s == null)
            return "";
        return s.replaceAll("^/+", "").replaceAll("/+$", "");
    }
}
