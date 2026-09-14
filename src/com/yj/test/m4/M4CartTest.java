package com.yj.test.m4;

import com.yj.bean.Cart;
import com.yj.bean.CartItem;
import com.yj.test.common.HttpTestClient;
import com.yj.test.common.HttpTestClient.Response;
import org.junit.BeforeClass;
import org.junit.Test;

import java.math.BigDecimal;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

/**
 * M4 购物车模块 自动化测试。
 *
 * <p>用例编号 BS-UT-061 ~ BS-IT-068，与《测试用例清单》一一对应。
 *
 * <p>其中 061、062 为纯单元测试（直接构造 {@link Cart}，不依赖容器与网络）；
 * 063 ~ 068 为集成测试（通过 HTTP 打真实的 Servlet，购物车状态保存在会话中，
 * 因此同一 {@link HttpTestClient} 实例内的多次请求共享一个会话）。
 *
 * <p>本模块的 5 个缺陷（BS-IT-063 ~ 067）已在修复中处理，用例现作为**回归测试**存在：
 * 断言写的是修复后的正确行为，若缺陷回归会立即失败。
 *
 * @author 王尚清
 */
public class M4CartTest {

    private static final String BASE = "http://localhost:8080/Book";

    /** 与清单一致：解忧杂货店，单价 27.20。 */
    private static final BigDecimal PRICE = new BigDecimal("27.20");

    private static HttpTestClient http;

    @BeforeClass
    public static void setUpClass() {
        http = new HttpTestClient(BASE);
    }

    // ============================================================ 单元测试

    /** BS-UT-061 同一本书重复加购时数量累加且金额正确（等价类）。 */
    @Test
    public void bsUt061_repeatAddAccumulatesCountAndPrice() {
        Cart cart = new Cart();
        cart.addItem(item(1, PRICE, "27.20"));
        cart.addItem(item(1, PRICE, "27.20"));   // 同一本书再次加购

        assertEquals("重复加购同一本书，购物车条目数应仍为 1（不应新建行）",
                1, cart.getItems().size());
        assertEquals("数量应累加为 2", 2, cart.getTotalCount().intValue());
        assertEquals("总金额应为 27.20 x 2 = 54.40",
                0, new BigDecimal("54.40").compareTo(cart.getTotalPrice()));
    }

    /** BS-UT-062 购物车总金额使用 BigDecimal 精确计算（等价类-数据精度）。 */
    @Test
    public void bsUt062_totalPriceUsesBigDecimalPrecision() {
        Cart cart = new Cart();
        cart.addItem(item(1, PRICE, "27.20"));

        // updateCount 会走 price.multiply(count) 这条真实的金额计算路径
        cart.updateCount(1, 3);

        assertEquals("27.20 x 3 应精确等于 81.60，不应出现浮点误差",
                "81.60", cart.getTotalPrice().toPlainString());
    }

    // ============================================================ 集成测试

    /** BS-IT-063 加购不存在的图书 ID 应被友好处理，不得返回 500（边界值）。 */
    @Test
    public void bsIt063_addNonexistentBookIsHandled() throws Exception {
        HttpTestClient c = new HttpTestClient(BASE);

        Response r = c.get("/cartServlet?action=ajaxAddItem&id=99999");

        assertTrue(
                "加购不存在的图书 ID 时应友好提示，不得返回 500。实际 " + r
                        + "，响应片段: " + snippet(r.body),
                r.status != 500);
    }

    /** BS-IT-064 加购非数字 ID 应被友好处理，不得返回 500（边界值）。 */
    @Test
    public void bsIt064_addNonNumericIdIsHandled() throws Exception {
        HttpTestClient c = new HttpTestClient(BASE);

        Response r = c.get("/cartServlet?action=ajaxAddItem&id=abc");

        assertTrue(
                "加购非数字 ID 时应提示参数非法，不得返回 500。实际 " + r
                        + "，响应片段: " + snippet(r.body),
                r.status != 500);
    }

    /** BS-IT-065 修改商品数量为负数应被拒绝（边界值）。 */
    @Test
    public void bsIt065_negativeCountIsRejected() throws Exception {
        HttpTestClient c = new HttpTestClient(BASE);
        c.get("/cartServlet?action=ajaxAddItem&id=1");   // 总数量 1
        c.get("/cartServlet?action=ajaxAddItem&id=2");   // 总数量 2

        c.get("/cartServlet?action=updateCount&id=1&count=-5");   // 非法数量，应被拒绝

        Response r = c.get("/cartServlet?action=ajaxAddItem&id=3");
        int total = totalCount(r.body);

        assertEquals(
                "负数数量应被拒绝并保留原值：id=1 的数量应仍为 1，再加购 1 件后总数量为 3。"
                        + "实际为 " + total + "（修复前曾被改成 -5，总数量变成 -3）。响应: " + r.body,
                3, total);
    }

    /** BS-IT-066 修改数量为非数字应被拒绝并保留原值（边界值）。 */
    @Test
    public void bsIt066_nonNumericCountIsRejected() throws Exception {
        HttpTestClient c = new HttpTestClient(BASE);
        c.get("/cartServlet?action=ajaxAddItem&id=1");
        c.get("/cartServlet?action=updateCount&id=1&count=5");     // 数量改为 5，总数量 5

        c.get("/cartServlet?action=updateCount&id=1&count=abc");   // 非法值，应被拒绝

        Response r = c.get("/cartServlet?action=ajaxAddItem&id=2");
        int total = totalCount(r.body);

        assertEquals(
                "数量参数非法时应保留原值 5：总数量应为 6（5+1）。"
                        + "实际为 " + total + "（修复前 abc 被静默改成 1，总数量变成 2）。响应: " + r.body,
                6, total);
    }

    /** BS-IT-067 空购物车时删除商品仍应给出响应（边界值）。 */
    @Test
    public void bsIt067_emptyCartDeleteStillResponds() throws Exception {
        HttpTestClient c = new HttpTestClient(BASE);
        c.newSession();     // 全新会话，会话中不存在 cart

        Response r = c.get("/cartServlet?action=deleteItem&id=1");

        assertTrue(
                "空购物车执行删除时也必须给出响应（跳转或提示），实际 " + r
                        + "：修复前既不跳转也不写出任何内容，客户端会停在空白页。",
                r.status == 302 || r.length() > 0);
    }

    /** BS-IT-068 清空购物车后会话中的 cart 被移除（场景法）。 */
    @Test
    public void bsIt068_clearItemEmptiesCart() throws Exception {
        HttpTestClient c = new HttpTestClient(BASE);
        c.get("/cartServlet?action=ajaxAddItem&id=1");
        Response before = c.get("/cartServlet?action=ajaxAddItem&id=2");
        assertEquals("前置条件：清空前购物车应有 2 件商品", 2, totalCount(before.body));

        c.get("/cartServlet?action=clearItem");

        Response after = c.get("/cartServlet?action=ajaxAddItem&id=3");
        assertEquals("清空后购物车应重新开始计数（仅剩本次加购的 1 件）",
                1, totalCount(after.body));
    }

    // ============================================================ 工具

    private static CartItem item(int id, BigDecimal price, String total) {
        return new CartItem(id, "测试图书" + id, 1, price, new BigDecimal(total));
    }

    /** 从 ajaxAddItem 返回的 JSON 中取出 totalCount。 */
    static int totalCount(String json) {
        if (json == null) {
            return Integer.MIN_VALUE;
        }
        Matcher m = Pattern.compile("\"totalCount\"\\s*:\\s*(-?\\d+)").matcher(json);
        return m.find() ? Integer.parseInt(m.group(1)) : Integer.MIN_VALUE;
    }

    private static String snippet(String body) {
        if (body == null) {
            return "<空>";
        }
        String s = body.replaceAll("\\s+", " ").trim();
        return s.length() > 120 ? s.substring(0, 120) + "..." : s;
    }
}
