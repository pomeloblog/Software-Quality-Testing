import org.junit.AfterClass;
import org.junit.FixMethodOrder;
import org.junit.Test;
import org.junit.runners.MethodSorters;

import static org.junit.Assert.*;

/** 按用例编号升序执行，输出与 Excel 清单顺序一致。 */
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
/**
 * M3 图书浏览与检索 —— 丙负责，共 6 条用例（BS-IT-041 ~ BS-IT-046）。
 * 方法覆盖：边界值法（041/042/045）、等价类法（043/045）、缺陷定向验证（044/046）。
 * 断言口径：预期结果一律按"系统应有的正确行为"书写，断言失败即视为用例 NG（缺陷证据）。
 */
public class M3BookBrowseTest {

    /** 统计页面上展示的图书条数（首页列表每本书渲染一个 "书名:" 标记）。 */
    private static int countBookItems(String body) {
        int n = 0, idx = 0;
        while ((idx = body.indexOf("书名:", idx)) >= 0) { n++; idx += 3; }
        return n;
    }

    /** BS-IT-041 首页分页 pageNo 边界容错（边界值法）。预期：越界/非法值回退到有效页，HTTP 200。 */
    @Test
    public void test041_pageNoBoundary() throws Exception {
        HttpTestUtil http = new HttpTestUtil();
        long total = HttpTestUtil.queryLong("select count(*) from t_book");
        int pageSize = 4;
        int lastPage = (int) ((total + pageSize - 1) / pageSize);

        // 合法值：第 1 页显示 4 本
        HttpTestUtil.Response r1 = http.get("/client/bookServlet?action=page&pageNo=1");
        assertEquals(200, r1.code);
        assertEquals(pageSize, countBookItems(r1.body));

        // 边界：pageNo=0 / -1 应回退到第 1 页
        HttpTestUtil.Response r0 = http.get("/client/bookServlet?action=page&pageNo=0");
        assertEquals("pageNo=0 应回退第1页且返回200", 200, r0.code);
        assertEquals("pageNo=0 应显示第1页的4本书", pageSize, countBookItems(r0.body));

        HttpTestUtil.Response rn = http.get("/client/bookServlet?action=page&pageNo=-1");
        assertEquals("pageNo=-1 应回退第1页且返回200", 200, rn.code);

        // 边界：pageNo 超大应回退到末页（末页仅剩 total%pageSize 本）
        HttpTestUtil.Response rx = http.get("/client/bookServlet?action=page&pageNo=99999");
        assertEquals("pageNo=99999 应回退末页且返回200", 200, rx.code);
        int lastPageCount = (int) (total % pageSize == 0 ? pageSize : total % pageSize);
        assertEquals("超大pageNo应回退末页", lastPageCount, countBookItems(rx.body));

        // 非法：非数字应取默认值 1
        HttpTestUtil.Response ra = http.get("/client/bookServlet?action=page&pageNo=abc");
        assertEquals("非数字pageNo应取默认1且返回200", 200, ra.code);
        assertEquals("非数字pageNo应显示第1页", pageSize, countBookItems(ra.body));
    }

    /**
     * BS-IT-042 分页 pageSize 边界（边界值法）。
     * 预期：pageSize 为 0 / 负数 / 非数字时应容错为默认值 4，不应出现服务器错误。
     */
    @Test
    public void test042_pageSizeBoundary() throws Exception {
        HttpTestUtil http = new HttpTestUtil();

        // 合法值：pageSize=6 正常返回 6 本
        HttpTestUtil.Response r6 = http.get("/client/bookServlet?action=page&pageNo=1&pageSize=6");
        assertEquals(200, r6.code);
        assertEquals(6, countBookItems(r6.body));

        // 非数字：容错为默认 4
        HttpTestUtil.Response ra = http.get("/client/bookServlet?action=page&pageNo=1&pageSize=abc");
        assertEquals("非数字pageSize应取默认值4", 200, ra.code);
        assertEquals(4, countBookItems(ra.body));

        // 边界：pageSize=0 应容错，不应 500（除零风险点）
        HttpTestUtil.Response r0 = http.get("/client/bookServlet?action=page&pageNo=1&pageSize=0");
        assertNotEquals("pageSize=0 不应导致服务器错误(500)", 500, r0.code);
        assertEquals("pageSize=0 应容错为默认每页4条", 4, countBookItems(r0.body));

        // 边界：pageSize=-1 应容错，不应 500（SQL LIMIT 负数风险点）
        HttpTestUtil.Response rn = http.get("/client/bookServlet?action=page&pageNo=1&pageSize=-1");
        assertNotEquals("pageSize=-1 不应导致服务器错误(500)", 500, rn.code);
    }

    /** BS-IT-043 书名/作者搜索等价类（等价类法：合法词 / 不存在词 / 空串 / 特殊字符）。 */
    @Test
    public void test043_searchEquivalence() throws Exception {
        HttpTestUtil http = new HttpTestUtil();

        // 合法等价类：命中书名
        HttpTestUtil.Response r1 = http.get("/client/bookServlet?action=pageByNameOrAuthor&nameorauthor=" + HttpTestUtil.enc("三体"));
        assertEquals(200, r1.code);
        long expect = HttpTestUtil.queryLong("select count(*) from t_book where name like ? or author like ?",
                "%三体%", "%三体%");
        assertTrue("搜索'三体'应命中" + expect + "本，页面至少展示1本", countBookItems(r1.body) >= 1);
        assertTrue(r1.body.contains("三体"));

        // 合法等价类：命中作者（东野圭吾 → 解忧杂货店）
        HttpTestUtil.Response r2 = http.get("/client/bookServlet?action=pageByNameOrAuthor&nameorauthor=" + HttpTestUtil.enc("东野圭吾"));
        assertEquals(200, r2.code);
        assertTrue("按作者搜索应返回其作品", r2.body.contains("解忧杂货店"));

        // 无效等价类：不存在的词 → 空结果，页面不报错
        HttpTestUtil.Response r3 = http.get("/client/bookServlet?action=pageByNameOrAuthor&nameorauthor=" + HttpTestUtil.enc("zzz不存在zzz"));
        assertEquals(200, r3.code);
        assertEquals("不存在的搜索词应返回空列表", 0, countBookItems(r3.body));

        // 特殊字符：单引号（SQL 注入尝试）→ 参数化查询应安全返回空/正常，不得 500
        HttpTestUtil.Response r4 = http.get("/client/bookServlet?action=pageByNameOrAuthor&nameorauthor=" + HttpTestUtil.enc("' or 1=1 --"));
        assertEquals("含单引号的搜索词不应导致服务器错误", 200, r4.code);
    }

    /**
     * BS-IT-044 搜索翻页链接对特殊字符的兼容（缺陷定向：URL 未编码）。
     * 预期：搜索词含空格/& 时，翻页链接应对参数值做 URL 编码，保证点击"下一页"后搜索条件不丢失。
     */
    @Test
    public void test044_searchPaginationEncoding() throws Exception {
        HttpTestUtil http = new HttpTestUtil();

        // 搜索词含空格："三 体"
        HttpTestUtil.Response r1 = http.get("/client/bookServlet?action=pageByNameOrAuthor&nameorauthor=" + HttpTestUtil.enc("三 体"));
        assertEquals(200, r1.code);
        // 若未编码，页面翻页链接会形如 &nameorauthor=三 体&pageNo=2（空格裸露，浏览器点击后参数被截断）
        assertFalse("翻页链接中的搜索词空格应被URL编码，不应裸露：发现 &nameorauthor=三 体&",
                r1.body.contains("nameorauthor=三 体&"));

        // 搜索词含 &："a&b"（点击翻页后 b 会被当作独立参数丢失）
        HttpTestUtil.Response r2 = http.get("/client/bookServlet?action=pageByNameOrAuthor&nameorauthor=" + HttpTestUtil.enc("a&b"));
        assertEquals(200, r2.code);
        assertFalse("翻页链接中的搜索词含&时必须编码，否则参数边界被破坏：发现 &nameorauthor=a&b&",
                r2.body.contains("nameorauthor=a&b&"));
    }

    /**
     * BS-IT-045 价格区间筛选边界（边界值法 + 等价类）。
     * 预期：区间合法时结果与库中数据一致；min>max、min=max、小数价格、非法值均有明确定义的行为。
     */
    @Test
    public void test045_priceFilterBoundary() throws Exception {
        HttpTestUtil http = new HttpTestUtil();

        // 合法区间 [10,30]：页面第 1 页条数 = min(4, 库中 10~30 元书数)
        long expect1030 = HttpTestUtil.queryLong("select count(*) from t_book where price between 10 and 30");
        assertTrue("预备：库中应存在10~30元的书", expect1030 > 0);
        HttpTestUtil.Response r1 = http.get("/client/bookServlet?action=pageByPrice&min=10&max=30");
        assertEquals(200, r1.code);
        assertEquals((int) Math.min(4, expect1030), countBookItems(r1.body));

        // 边界 min=max：恰好匹配该价格的书（库中存在 56.5 元的《三体》）
        long expect565 = HttpTestUtil.queryLong("select count(*) from t_book where price = 56.5");
        HttpTestUtil.Response r2 = http.get("/client/bookServlet?action=pageByPrice&min=57&max=57");
        long expect57 = HttpTestUtil.queryLong("select count(*) from t_book where price = 57");
        assertEquals(200, r2.code);
        assertEquals("min=max=57 应精确匹配" + expect57 + "本书", (int) expect57, countBookItems(r2.body));

        // 边界 min>max：系统应给出提示或归一化处理，而不是静默返回空页无任何说明
        HttpTestUtil.Response r3 = http.get("/client/bookServlet?action=pageByPrice&min=100&max=10");
        assertEquals("min>max 不应导致服务器错误", 200, r3.code);
        assertTrue("min>max 时应给出可见的错误/提示信息（或自动交换区间）",
                countBookItems(r3.body) == 0 && r3.body.contains("提示") || r3.body.contains("区间") || r3.body.contains("错误"));

        // 小数价格边界：库中价格是 DECIMAL（如 56.5），传小数应正常筛选
        HttpTestUtil.Response r4 = http.get("/client/bookServlet?action=pageByPrice&min=56.5&max=56.5");
        assertEquals("小数价格参数不应导致服务器错误", 200, r4.code);
        assertEquals("min=max=56.5 应精确匹配价格为56.5的书（当前库中有" + expect565 + "本）",
                (int) expect565, countBookItems(r4.body));
    }

    /**
     * BS-IT-046 销量榜单排序与完整性（缺陷定向：LIMIT 跳过第一名 + 排序正确性）。
     * 预期：榜单按销量降序展示，且必须包含销量最高的书（第一名）。
     * 造数：把 id=1 的书销量改为 999（全库最高），结束后还原。
     */
    @Test
    public void test046_salesRanking() throws Exception {
        HttpTestUtil http = new HttpTestUtil();
        String origSales = HttpTestUtil.queryString("select sales from t_book where id=1");
        String topName = HttpTestUtil.queryString("select name from t_book where id=1");
        try {
            HttpTestUtil.exec("update t_book set sales=999 where id=1");

            HttpTestUtil.Response r = http.get("/client/bookServlet?action=pageOrder");
            assertEquals(200, r.code);
            assertTrue("销量榜单必须包含销量第一名《" + topName + "》",
                    r.body.contains(topName));
        } finally {
            HttpTestUtil.exec("update t_book set sales=? where id=1", Integer.valueOf(origSales));
        }
    }

    /** 兜底还原：防止 046 异常路径漏还原。 */
    @AfterClass
    public static void restore() throws Exception {
        HttpTestUtil.exec("update t_book set sales=100 where id=1 and sales=999");
    }
}
