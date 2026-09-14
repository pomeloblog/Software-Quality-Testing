package com.yj.test.common;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 集成测试用极简 HTTP 客户端。
 *
 * <p>自带 Cookie（JSESSIONID）管理，因此同一个实例的多次请求共享一个会话——
 * 购物车存在 session 中，必须靠它维持会话状态。不自动跟随 302 跳转，
 * 便于直接断言服务端返回的状态码。
 *
 * @author 王尚清
 */
public class HttpTestClient {

    /** 单次响应。 */
    public static class Response {
        public final int status;
        public final String body;

        Response(int status, String body) {
            this.status = status;
            this.body = body;
        }

        public boolean contains(String text) {
            return body != null && body.contains(text);
        }

        public int length() {
            return body == null ? 0 : body.length();
        }

        @Override
        public String toString() {
            return "HTTP " + status + " / " + length() + " bytes";
        }
    }

    private final String base;
    private final Map<String, String> cookies = new LinkedHashMap<String, String>();

    public HttpTestClient(String base) {
        this.base = base.endsWith("/") ? base.substring(0, base.length() - 1) : base;
    }

    public String getBase() {
        return base;
    }

    /** 丢弃全部 Cookie，相当于换一个全新会话（购物车为空）。 */
    public void newSession() {
        cookies.clear();
    }

    /** 探测应用是否已启动。 */
    public boolean serverReachable() {
        try {
            get("/");
            return true;
        } catch (IOException e) {
            return false;
        }
    }

    public Response get(String path) throws IOException {
        return send("GET", path, null);
    }

    public Response post(String path, Map<String, String> form) throws IOException {
        return send("POST", path, form);
    }

    /** 按 name=value 交替传入，快速构造表单参数。 */
    public static Map<String, String> form(String... kv) {
        Map<String, String> m = new LinkedHashMap<String, String>();
        for (int i = 0; i + 1 < kv.length; i += 2) {
            m.put(kv[i], kv[i + 1]);
        }
        return m;
    }

    private Response send(String method, String path, Map<String, String> form) throws IOException {
        HttpURLConnection conn = (HttpURLConnection) new URL(base + path).openConnection();
        conn.setRequestMethod(method);
        conn.setConnectTimeout(5000);
        conn.setReadTimeout(15000);
        conn.setInstanceFollowRedirects(false);
        conn.setRequestProperty("User-Agent", "BookStore-AutoTest/1.0");

        if (!cookies.isEmpty()) {
            StringBuilder sb = new StringBuilder();
            for (Map.Entry<String, String> e : cookies.entrySet()) {
                if (sb.length() > 0) {
                    sb.append("; ");
                }
                sb.append(e.getKey()).append('=').append(e.getValue());
            }
            conn.setRequestProperty("Cookie", sb.toString());
        }

        if (form != null) {
            conn.setDoOutput(true);
            conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
            OutputStream os = conn.getOutputStream();
            os.write(encode(form).getBytes("UTF-8"));
            os.close();
        }

        int status = conn.getResponseCode();
        collectCookies(conn);

        InputStream is = status >= 400 ? conn.getErrorStream() : conn.getInputStream();
        String body = is == null ? "" : read(is);
        conn.disconnect();
        return new Response(status, body);
    }

    private void collectCookies(HttpURLConnection conn) {
        for (Map.Entry<String, List<String>> header : conn.getHeaderFields().entrySet()) {
            if (header.getKey() == null || !"set-cookie".equalsIgnoreCase(header.getKey())) {
                continue;
            }
            for (String raw : header.getValue()) {
                int eq = raw.indexOf('=');
                int semi = raw.indexOf(';');
                if (eq > 0) {
                    String name = raw.substring(0, eq).trim();
                    String value = (semi > eq ? raw.substring(eq + 1, semi) : raw.substring(eq + 1)).trim();
                    cookies.put(name, value);
                }
            }
        }
    }

    private static String encode(Map<String, String> form) {
        StringBuilder sb = new StringBuilder();
        for (Map.Entry<String, String> e : form.entrySet()) {
            if (sb.length() > 0) {
                sb.append('&');
            }
            sb.append(urlEncode(e.getKey())).append('=').append(urlEncode(e.getValue()));
        }
        return sb.toString();
    }

    private static String urlEncode(String s) {
        try {
            return URLEncoder.encode(s == null ? "" : s, "UTF-8");
        } catch (UnsupportedEncodingException e) {
            throw new IllegalStateException(e);
        }
    }

    private static String read(InputStream is) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] buf = new byte[4096];
        int n;
        while ((n = is.read(buf)) != -1) {
            bos.write(buf, 0, n);
        }
        is.close();
        return new String(bos.toByteArray(), "UTF-8");
    }
}
