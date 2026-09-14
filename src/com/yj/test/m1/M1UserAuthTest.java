package com.yj.test.m1;

import com.yj.filter.TransactionFilter;
import com.yj.test.common.HttpTestClient;
import com.yj.test.common.HttpTestClient.Response;
import com.yj.test.common.MockHttp;
import com.yj.test.common.TestDb;
import com.yj.web.BaseServlet;
import com.yj.web.UserServlet;
import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

import javax.servlet.FilterChain;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

/**
 * M1 用户认证模块 自动化测试。
 *
 * <p>用例编号 BS-IT-001 ~ BS-IT-007、BS-UT-006，与《测试用例清单》一一对应。
 * 方法名规则：{@code bsIt00X_描述} / {@code bsUt00X_描述}，由 {@code TestRunner} 还原成用例编号。
 *
 * <p><b>为什么登录用例用桩而不是走 HTTP：</b>登录已在修复中补上了验证码校验，
 * 而自动化脚本无法识别验证码图片（这正是验证码的设计目的）。
 * 因此登录相关的用例改用 {@link MockHttp} 向会话注入验证码后直接驱动 Servlet，
 * 走的是「Servlet → Service → DAO → MySQL」的完整链路外加真实的 {@link TransactionFilter}，
 * 仅替换掉容器的传输层。
 *
 * <p>注册、验证码等仍走真实 HTTP 的用例保持不变。
 *
 * @author 王尚清
 */
public class M1UserAuthTest {

    private static final String BASE = "http://localhost:8080/Book";

    /** 桩注入到会话中的验证码。 */
    private static final String CAPTCHA = "TEST";

    private static HttpTestClient http;

    @BeforeClass
    public static void setUpClass() {
        http = new HttpTestClient(BASE);
    }

    @Before
    public void setUp() {
        TestDb.deleteUser("probe01");
        TestDb.deleteUser("repwd01");
    }

    @After
    public void tearDown() {
        TestDb.deleteUser("repwd01");
    }

    // ------------------------------------------------------------------ 登录

    /** BS-IT-001 正确用户名和密码登录成功（等价类-有效）。 */
    @Test
    public void bsIt001_loginSucceedsWithCorrectPassword() throws Exception {
        MockHttp mock = new MockHttp(HttpTestClient.form(
                "username", "admin",
                "password", "admin",
                "code", CAPTCHA,
                "action", "login")).withCaptcha(CAPTCHA);

        invokeThroughFilter(mock);

        assertEquals("正确账号密码应登录成功并转发到成功页，实际: " + mock.describe(),
                "/pages/user/login_success.jsp", mock.forwardedTo);
        assertNotNull("登录成功后会话中应写入 user 对象", mock.sessionAttrs.get("user"));
    }

    /** BS-IT-002 密码错误时登录失败（等价类-无效）。 */
    @Test
    public void bsIt002_loginFailsWithWrongPassword() throws Exception {
        MockHttp mock = new MockHttp(HttpTestClient.form(
                "username", "admin",
                "password", "wrongpwd",
                "code", CAPTCHA,
                "action", "login")).withCaptcha(CAPTCHA);

        invokeThroughFilter(mock);

        assertEquals("密码错误应回到登录页，实际: " + mock.describe(),
                "/pages/user/login.jsp", mock.forwardedTo);
        assertEquals("应提示用户名或密码错误", "用户名或密码错误!", mock.requestAttrs.get("msg"));
        assertNull("登录失败不应写入会话", mock.sessionAttrs.get("user"));
    }

    /** BS-IT-003 用户名不存在时登录失败（等价类-无效）。 */
    @Test
    public void bsIt003_loginFailsWithUnknownUsername() throws Exception {
        MockHttp mock = new MockHttp(HttpTestClient.form(
                "username", "nobody_xyz",
                "password", "whatever",
                "code", CAPTCHA,
                "action", "login")).withCaptcha(CAPTCHA);

        invokeThroughFilter(mock);

        assertEquals("用户不存在应回到登录页，实际: " + mock.describe(),
                "/pages/user/login.jsp", mock.forwardedTo);
        assertEquals("应提示用户名或密码错误", "用户名或密码错误!", mock.requestAttrs.get("msg"));
        assertNull("登录失败不应写入会话", mock.sessionAttrs.get("user"));
    }

    /**
     * BS-IT-004 登录请求不携带验证码应被拒绝（场景法）。
     *
     * <p>回归验证点：修复前登录完全未校验验证码，本用例判定为 NG（缺陷 BS-IT-004）；
     * 修复后应通过。
     */
    @Test
    public void bsIt004_loginRejectedWithoutCaptcha() throws Exception {
        HttpTestClient fresh = new HttpTestClient(BASE);
        fresh.newSession();

        Response r = fresh.post("/userServlet", HttpTestClient.form(
                "username", "admin",
                "password", "admin",
                "action", "login"));

        assertFalse(
                "登录必须校验验证码：未携带 code 参数（会话中也没有验证码）时不应登录成功。实际返回 " + r,
                r.contains("登录成功"));
        assertTrue("应提示验证码错误", r.contains("验证码错误"));
    }

    // ------------------------------------------------------------------ 注册

    /** BS-IT-005 注册时验证码缺失被拒绝且不落库（场景法）。 */
    @Test
    public void bsIt005_registRejectedWithoutCaptcha() throws Exception {
        HttpTestClient fresh = new HttpTestClient(BASE);
        fresh.newSession();

        Response r = fresh.post("/userServlet", HttpTestClient.form(
                "username", "probe01",
                "password", "123456",
                "email", "probe01@test.com",
                "action", "regist"));

        assertEquals(200, r.status);
        assertTrue("缺少验证码时应提示「验证码错误」", r.contains("验证码错误"));
        assertEquals("验证码缺失时用户不应被写入数据库", 0, TestDb.countUser("probe01"));
    }

    /**
     * BS-UT-006 注册时两次密码不一致应被拒绝（等价类-无效）。
     *
     * <p>回归验证点：修复前 regist() 取出 repwd 后从未比对，两次密码不同也能注册成功
     * （缺陷 BS-UT-006）；修复后应在落库前拦截。
     */
    @Test
    public void bsUt006_mismatchedPasswordIsRejected() throws Exception {
        MockHttp mock = new MockHttp(HttpTestClient.form(
                "username", "repwd01",
                "password", "123456",
                "repwd", "654321",          // 与 password 故意不一致
                "email", "repwd01@test.com",
                "code", CAPTCHA,
                "action", "regist")).withCaptcha(CAPTCHA);

        invokeThroughFilter(mock);

        assertEquals(
                "两次密码不一致时应拒绝注册并留在注册页，实际: " + mock.describe(),
                "/pages/user/regist.jsp", mock.forwardedTo);
        assertEquals("应提示两次密码不一致", "两次输入的密码不一致！", mock.requestAttrs.get("msg"));
        assertEquals("密码不一致时用户不应被写入数据库", 0, TestDb.countUser("repwd01"));
    }

    // ------------------------------------------------------- 用户名查重（辅助）

    /** BS-IT-007 用户名查重接口返回值正确（等价类）。 */
    @Test
    public void bsIt007_ajaxExistsUsername() throws Exception {
        Response exists = http.get("/userServlet?action=ajaxExistsusername&username=admin");
        assertEquals(200, exists.status);
        assertTrue("已存在的用户名应返回 exitsUsername=true，实际: " + exists.body,
                exists.body.contains("true"));

        Response notExists = http.get("/userServlet?action=ajaxExistsusername&username=notexist_xyz");
        assertEquals(200, notExists.status);
        assertTrue("不存在的用户名应返回 exitsUsername=false，实际: " + notExists.body,
                notExists.body.contains("false"));
    }

    // ------------------------------------------------------------------ 工具

    /**
     * 以生产一致的方式驱动 Servlet：外层套真实的 {@link TransactionFilter}。
     *
     * <p>被测代码把事务边界交给了 TransactionFilter（负责 commit / rollback），
     * {@code BaseServlet} 内部的 DAO 拿到的是 autoCommit=false 的连接。
     * 若绕过过滤器直接调用，写操作既不会提交（断言永远看不到数据），
     * 又会因未提交的 INSERT 持有行锁而拖住后续清理语句。
     */
    private static void invokeThroughFilter(MockHttp mock) throws Exception {
        final UserServlet servlet = new UserServlet();
        final Method doPost = BaseServlet.class.getDeclaredMethod(
                "doPost", HttpServletRequest.class, HttpServletResponse.class);
        doPost.setAccessible(true);

        new TransactionFilter().doFilter(mock.request(), mock.response(), new FilterChain() {
            @Override
            public void doFilter(ServletRequest request, ServletResponse response)
                    throws IOException, ServletException {
                try {
                    doPost.invoke(servlet, request, response);
                } catch (InvocationTargetException e) {
                    throw new ServletException(e.getCause());
                } catch (IllegalAccessException e) {
                    throw new ServletException(e);
                }
            }
        });
    }
}
