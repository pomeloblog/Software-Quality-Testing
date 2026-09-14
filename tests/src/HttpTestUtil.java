import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.nio.charset.StandardCharsets;

/**
 * 测试基础设施：HTTP 请求（带会话）+ JDBC 直连数据库断言。
 * 被测系统：http://localhost:8080/Book
 * 数据库：book@localhost:3306（与被测应用同库，用于造数与断言）
 */
public final class HttpTestUtil {

    public static final String SERVER_ROOT = "http://localhost:8080";
    public static final String BASE = SERVER_ROOT + "/Book";
    public static final String DB_URL = "jdbc:mysql://localhost:3306/book?characterEncoding=utf8&useSSL=false";
    public static final String DB_USER = "bookstore";
    public static final String DB_PASS = "123456";

    private String jsessionId;
    /** 仅 login() 期间为 true：重定向跟随产生的临时会话不得覆盖登录态。 */
    private boolean captureSession;

    /** 发 GET 请求（手动跟随重定向，全程保持会话）。 */
    public Response get(String pathAndQuery) throws Exception {
        return send("GET", pathAndQuery, null);
    }

    /** 发 POST application/x-www-form-urlencoded 请求。kv 形如 "action=login&username=admin"。 */
    public Response post(String pathAndQuery, String formBody) throws Exception {
        return send("POST", pathAndQuery, formBody);
    }

    /** 发 GET 且不携带会话（用于未登录场景）。 */
    public static Response getNoSession(String pathAndQuery) throws Exception {
        return new HttpTestUtil().send("GET", pathAndQuery, null);
    }

    /** 用指定账号登录，成功返回 true 并记住 JSESSIONID。 */
    public boolean login(String username, String password) throws Exception {
        captureSession = true;
        try {
            Response r = send("POST", "/userServlet?action=login",
                    "action=login&username=" + enc(username) + "&password=" + enc(password));
            return r.code == 200 && r.body.contains("登录成功");
        } finally {
            captureSession = false;
        }
    }

    /**
     * 手动跟随重定向：每跳都重发原 Cookie（JDK 自动跟随不传 Cookie，会把
     * 需登录的 302 跟成登录页并返回新 JSESSIONID，污染会话——已踩坑）。
     * 302 之后的跳转按浏览器惯例降级为 GET。
     */
    private Response send(String method, String pathAndQuery, String formBody) throws Exception {
        String url = BASE + pathAndQuery;
        String curMethod = method;
        String curBody = formBody;
        for (int hop = 0; hop < 5; hop++) {
            HttpURLConnection conn = (HttpURLConnection) new URL(url).openConnection();
            conn.setRequestMethod(curMethod);
            conn.setInstanceFollowRedirects(false);
            conn.setConnectTimeout(10000);
            conn.setReadTimeout(20000);
            if (jsessionId != null) {
                conn.setRequestProperty("Cookie", "JSESSIONID=" + jsessionId);
            }
            // 模拟浏览器行为带 Referer（部分 Servlet 依赖它做回跳，如 sendOrder）
            conn.setRequestProperty("Referer", BASE + "/");
            if (curBody != null) {
                conn.setDoOutput(true);
                conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
                try (OutputStream os = conn.getOutputStream()) {
                    os.write(curBody.getBytes(StandardCharsets.UTF_8));
                }
            }
            int code = conn.getResponseCode();
            if (captureSession) {
                String cookie = conn.getHeaderField("Set-Cookie");
                if (cookie != null && cookie.contains("JSESSIONID=")) {
                    jsessionId = cookie.replaceAll(".*JSESSIONID=([^;]+).*", "$1");
                }
            }
            String location = conn.getHeaderField("Location");
            if ((code == 301 || code == 302 || code == 303) && location != null) {
                if (location.startsWith("http")) {
                    url = location;                       // 绝对 URL
                } else if (location.startsWith("/")) {
                    url = SERVER_ROOT + location;         // 服务器根绝对路径（sendRedirect(contextPath+..) 的情况）
                } else {
                    url = BASE + "/" + location;          // 相对路径
                }
                curMethod = "GET"; // 浏览器对 302 的行为：后续请求改为 GET
                curBody = null;
                continue;
            }
            InputStream in = (code >= 400) ? conn.getErrorStream() : conn.getInputStream();
            return new Response(code, readAll(in));
        }
        throw new IllegalStateException("too many redirects: " + url);
    }

    private static String readAll(InputStream in) throws Exception {
        if (in == null) return "";
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        byte[] b = new byte[8192];
        int n;
        while ((n = in.read(b)) > 0) buf.write(b, 0, n);
        return new String(buf.toByteArray(), StandardCharsets.UTF_8);
    }

    public static String enc(String v) {
        try {
            return URLEncoder.encode(v, "UTF-8");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static class Response {
        public final int code;
        public final String body;
        Response(int code, String body) { this.code = code; this.body = body; }
    }

    // ---------------- JDBC 断言工具 ----------------

    public static Connection db() throws SQLException {
        try {
            Class.forName("com.mysql.jdbc.Driver");
        } catch (ClassNotFoundException e) {
            throw new SQLException("mysql driver not on classpath", e);
        }
        return DriverManager.getConnection(DB_URL, DB_USER, DB_PASS);
    }

    /** 查询单个整数值，如 count(*) / sum(x)。NULL 返回 def。 */
    public static long queryLong(String sql, Object... params) throws SQLException {
        try (Connection c = db(); PreparedStatement ps = c.prepareStatement(sql)) {
            bind(ps, params);
            try (ResultSet rs = ps.executeQuery()) {
                rs.next();
                Object v = rs.getObject(1);
                return v == null ? 0L : ((Number) v).longValue();
            }
        }
    }

    /** 查询单个字符串值。 */
    public static String queryString(String sql, Object... params) throws SQLException {
        try (Connection c = db(); PreparedStatement ps = c.prepareStatement(sql)) {
            bind(ps, params);
            try (ResultSet rs = ps.executeQuery()) {
                rs.next();
                Object v = rs.getObject(1);
                return v == null ? null : v.toString();
            }
        }
    }

    /** 执行增删改，返回影响行数。 */
    public static int exec(String sql, Object... params) throws SQLException {
        try (Connection c = db(); PreparedStatement ps = c.prepareStatement(sql)) {
            bind(ps, params);
            return ps.executeUpdate();
        }
    }

    private static void bind(PreparedStatement ps, Object[] params) throws SQLException {
        for (int i = 0; i < params.length; i++) ps.setObject(i + 1, params[i]);
    }
}
