import org.junit.AfterClass;
import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.*;

/**
 * M6 后台管理 —— 丙负责，共 6 条用例（BS-IT-101 ~ BS-IT-106）。
 * 方法覆盖：场景法（101/103/106 完整管理流程）、等价类法（102/105 合法/非法输入）、边界值法（104 删除目标边界）。
 * 预置条件：Tomcat 运行中、book 库可用、admin/admin 为管理员账号。
 * 断言口径：预期结果按"系统应有的正确行为"书写，断言失败即视为用例 NG（缺陷证据）。
 */
public class M6ManagerTest {

    private HttpTestUtil http;

    @Before
    public void setUp() throws Exception {
        http = new HttpTestUtil();
        assertTrue("前置：管理员 admin/admin 应能登录", http.login("admin", "admin"));
    }

    /** BS-IT-101 管理员查看图书管理分页列表（场景法）。预期：登录后 200，列表分页正确。 */
    @Test
    public void test101_managerBookPage() throws Exception {
        HttpTestUtil.Response r = http.get("/manager/bookServlet?action=page&pageNo=1");
        assertEquals("管理员访问图书管理页应返回200", 200, r.code);
        assertTrue("图书管理页应包含分页导航", r.body.contains("首页"));

        // 第 1 页应展示库中前 4 本书（按 id 升序）
        String firstName = HttpTestUtil.queryString("select name from t_book order by id limit 1");
        assertTrue("图书管理页第1页应展示图书《" + firstName + "》", r.body.contains(firstName));
        long total = HttpTestUtil.queryLong("select count(*) from t_book");
        assertTrue("库中图书数应>0", total > 0);

        // 对照组：未登录（不带会话）访问后台图书管理应被 ManagerFilter 拦截（forward 到登录页，HTTP 200 但内容为登录表单）
        HttpTestUtil.Response anon = HttpTestUtil.getNoSession("/manager/bookServlet?action=page&pageNo=1");
        assertTrue("未登录访问应被拦截到登录页（响应应为登录表单）",
                anon.body.contains("登录") && !anon.body.contains(firstName));
    }

    /**
     * BS-IT-102 新增图书输入校验（等价类法）。
     * 预期：合法数据入库成功；空书名 / 负价格 / 非数字价格应被校验并拒绝（提示错误），不得直接入库。
     */
    @Test
    public void test102_addBookValidation() throws Exception {
        // 合法等价类：正常新增
        HttpTestUtil.Response ok = http.post("/manager/bookServlet?action=add",
                "action=add&name=丙测试书&author=CPTEST&price=35.50&sales=10&stock=10");
        assertEquals(200, ok.code);
        String price = HttpTestUtil.queryString("select price from t_book where author='CPTEST' and name='丙测试书' order by id desc limit 1");
        assertNotNull("合法图书应成功入库", price);
        assertEquals("35.50", price);

        // 无效等价类：空书名应被拒绝
        http.post("/manager/bookServlet?action=add",
                "action=add&name=&author=CPTEST&price=20.00&sales=5&stock=5");
        long emptyName = HttpTestUtil.queryLong("select count(*) from t_book where author='CPTEST' and (name='' or name is null)");
        assertEquals("空书名应被校验拒绝，库中不应出现无名图书", 0, emptyName);

        // 无效等价类：负价格应被拒绝
        http.post("/manager/bookServlet?action=add",
                "action=add&name=负价书&author=CPTEST&price=-5&sales=5&stock=5");
        long negPrice = HttpTestUtil.queryLong("select count(*) from t_book where author='CPTEST' and price<0");
        assertEquals("负价格应被校验拒绝", 0, negPrice);

        // 无效等价类：非数字价格应被拒绝或给出提示
        http.post("/manager/bookServlet?action=add",
                "action=add&name=坏价格书&author=CPTEST&price=abc&sales=5&stock=5");
        long badPrice = HttpTestUtil.queryLong("select count(*) from t_book where author='CPTEST' and name='坏价格书'");
        assertEquals("非数字价格应被校验拒绝，不应入库", 0, badPrice);
    }

    /** BS-IT-103 修改图书：回显正确 + 更新生效（场景法）。预期：getBook 回显原值，update 后库中数据一致。 */
    @Test
    public void test103_editBookEchoAndUpdate() throws Exception {
        // 选取 id=2 的书（《边城》）
        String origName = HttpTestUtil.queryString("select name from t_book where id=2");
        String origPrice = HttpTestUtil.queryString("select price from t_book where id=2");
        String origAuthor = HttpTestUtil.queryString("select author from t_book where id=2");
        int origSales = Integer.parseInt(HttpTestUtil.queryString("select sales from t_book where id=2"));
        int origStock = Integer.parseInt(HttpTestUtil.queryString("select stock from t_book where id=2"));

        try {
            // 回显：getBook 转发到 book_edit.jsp，表单 value 应为原值
            HttpTestUtil.Response echo = http.get("/manager/bookServlet?action=getBook&id=2");
            assertEquals(200, echo.code);
            assertTrue("编辑页应回显原书名《" + origName + "》", echo.body.contains("value=\"" + origName + "\""));

            // 更新：改价格
            HttpTestUtil.Response upd = http.post("/manager/bookServlet?action=update",
                    "action=update&id=2&name=" + HttpTestUtil.enc(origName) + "&price=99.99&author=" + HttpTestUtil.enc(origAuthor)
                            + "&sales=" + origSales + "&stock=" + origStock);
            assertEquals(200, upd.code);
            assertEquals("更新后库中价格应为99.99", "99.99", HttpTestUtil.queryString("select price from t_book where id=2"));
        } finally {
            // 还原
            HttpTestUtil.exec("update t_book set name=?, price=?, author=?, sales=?, stock=? where id=2",
                    origName, origPrice, origAuthor, origSales, origStock);
        }
    }

    /**
     * BS-IT-104 删除图书边界（边界值法：存在的 id / 不存在的 id / 非数字 id）。
     * 预期：存在的 id 删除成功；不存在的 id 给出友好提示；非数字 id 不应导致服务器 500。
     */
    @Test
    public void test104_deleteBookBoundary() throws Exception {
        // 造数：SQL 插入一本待删书
        HttpTestUtil.exec("insert into t_book(name,author,classification,price,sales,stock,imgpath) values('丙待删书','CPDEL','文学',9.90,1,1,'static/img/default.jpg')");
        long id = HttpTestUtil.queryLong("select id from t_book where author='CPDEL' order by id desc limit 1");

        // 边界内：存在的 id → 删除成功
        HttpTestUtil.Response r1 = http.get("/manager/bookServlet?action=delete&id=" + id);
        assertEquals("删除存在的图书后应正常跳回列表页", 200, r1.code);
        assertEquals("库中该图书应已被删除", 0,
                HttpTestUtil.queryLong("select count(*) from t_book where id=?", id));

        // 边界外：不存在的 id → 应有友好提示（而非静默成功）
        HttpTestUtil.Response r2 = http.get("/manager/bookServlet?action=delete&id=999999");
        assertEquals("删除不存在的id不应导致服务器错误", 200, r2.code);
        assertTrue("删除不存在的id应给出提示信息", r2.body.contains("不存在") || r2.body.contains("删除失败"));

        // 非法：非数字 id → 不应 500
        HttpTestUtil.Response r3 = http.get("/manager/bookServlet?action=delete&id=abc");
        assertNotEquals("非数字id不应导致服务器500错误", 500, r3.code);
    }

    /**
     * BS-IT-105 用户管理：新增重复用户名与正常新增（等价类法 + 场景法）。
     * 预期：重复用户名给出友好错误提示（不 500）；正常新增成功且可删除。
     */
    @Test
    public void test105_userManagement() throws Exception {
        long before = HttpTestUtil.queryLong("select count(*) from t_user");

        // 无效等价类：新增重复用户名 admin
        HttpTestUtil.Response dup = http.post("/manager/UserServlet?action=add",
                "action=add&username=admin&password=x&email=dup@test.com");
        assertTrue("新增重复用户名应给出友好提示页面（而非500）",
                dup.code == 200 && (dup.body.contains("已存在") || dup.body.contains("失败") || dup.body.contains("重复")));

        // 合法等价类：正常新增用户
        HttpTestUtil.Response ok = http.post("/manager/UserServlet?action=add",
                "action=add&username=cp_test_01&password=123456&email=cp@test.com");
        assertEquals(200, ok.code);
        assertEquals("正常新增用户后库中用户数+1", before + 1,
                HttpTestUtil.queryLong("select count(*) from t_user"));

        // 场景收尾：删除该用户
        long uid = HttpTestUtil.queryLong("select id from t_user where username='cp_test_01'");
        HttpTestUtil.Response del = http.get("/manager/UserServlet?action=delete&id=" + uid);
        assertEquals(200, del.code);
        assertEquals("删除后用户数还原", before, HttpTestUtil.queryLong("select count(*) from t_user"));
    }

    /**
     * BS-IT-106 总账单对账 + 发货状态流转（场景法）。
     * 预期：账单四项统计与直查数据库一致；发货后订单 status 由 0 变 1。
     */
    @Test
    public void test106_totalBillAndSendOrder() throws Exception {
        HttpTestUtil.Response r = http.get("/manager/orderServlet?action=showTotal");
        assertEquals(200, r.code);

        long userCnt = HttpTestUtil.queryLong("select count(*) from t_user");
        long orderCnt = HttpTestUtil.queryLong("select count(*) from t_order");
        long salesSum = HttpTestUtil.queryLong("select coalesce(sum(sales),0) from t_book");
        String moneySum = HttpTestUtil.queryString("select coalesce(sum(price*sales),0) from t_book");

        assertTrue("账单用户总数应与库一致(" + userCnt + ")，页面应包含 '" + userCnt + "人'",
                r.body.contains(userCnt + "人"));
        assertTrue("账单订单数应与库一致(" + orderCnt + ")，页面应包含 '" + orderCnt + "单'",
                r.body.contains(orderCnt + "单"));
        assertTrue("账单销售本数应与库一致(" + salesSum + ")，页面应包含 '" + salesSum + "本'",
                r.body.contains(salesSum + "本"));
        assertTrue("账单总收入应与库口径一致(" + moneySum + ")，页面应包含 '" + moneySum + "元'",
                r.body.contains(moneySum + "元"));

        // 发货状态流转：造一条 status=0 的订单 → sendOrder → status 应变为 1
        HttpTestUtil.exec("insert into t_order(order_id,create_time,price,status,user_id) values('CPTEST_ORDER_1',now(),10.00,0,1)");
        try {
            HttpTestUtil.Response send = http.get("/manager/orderServlet?action=sendOrder&orderId=CPTEST_ORDER_1");
            assertEquals("发货操作应正常完成", 200, send.code);
            assertEquals("发货后订单status应由0变为1", 1,
                    HttpTestUtil.queryLong("select status from t_order where order_id='CPTEST_ORDER_1'"));
        } finally {
            HttpTestUtil.exec("delete from t_order where order_id='CPTEST_ORDER_1'");
        }
    }

    /** 兜底清理：清掉本用例类可能遗留的测试数据。 */
    @AfterClass
    public static void cleanup() throws Exception {
        HttpTestUtil.exec("delete from t_book where author in ('CPTEST','CPDEL')");
        HttpTestUtil.exec("delete from t_user where username='cp_test_01'");
        HttpTestUtil.exec("delete from t_order where order_id='CPTEST_ORDER_1'");
    }
}
