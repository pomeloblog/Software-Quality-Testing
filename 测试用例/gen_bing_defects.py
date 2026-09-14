# -*- coding: utf-8 -*-
"""生成丙（M3+M6）的缺陷报告 Word 文档。用法：python gen_bing_defects.py"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "丙-缺陷报告-M3-M6.docx")

DEFECTS = [
    dict(
        id="BUG-M3-01", title="分页参数 pageSize 为 0 或负数时服务端 500 错误",
        module="M3 图书浏览与检索", sev="高", prio="高",
        case="BS-IT-042",
        symptom="GET /Book/client/bookServlet?action=page&pageNo=1&pageSize=0 返回 HTTP 500；pageSize=-1 同样 500。",
        steps="1. 打开 http://localhost:8080/Book/client/bookServlet?action=page&pageNo=1&pageSize=0\n2. 观察返回：HTTP 500 服务器错误页\n3. 将 pageSize 改为 -1 重复，同样 500",
        expect="非法 pageSize 应容错为默认每页 4 条，正常返回列表页。",
        actual="pageSize=0：BookServiceImpl.page() 中 pageTotalCount % pageSize 抛 ArithmeticException（除零）；pageSize=-1：pageTotal 为负数 → pageNo 归一为负 → LIMIT 起始/条数为负 → MySQL 语法错误。",
        root="src/com/yj/service/impl/BookServiceImpl.java 第 59、63-68 行：对 pageSize 未做合法性校验（<1 或非法时取默认值）。",
        fix="WebUtils.parseInt 已兜底非数字；需再对解析结果做 if (pageSize < 1) pageSize = Page.PAGE_SIZE; 的下界校验。",
        verify="修复后依次请求 pageSize=0/-1/abc，均应返回 200 且每页 4 条；用 tests\\run-tests.bat 复测 test042 应通过。"),
    dict(
        id="BUG-M3-02", title="搜索结果翻页链接未 URL 编码，含空格/& 的搜索条件丢失",
        module="M3 图书浏览与检索", sev="高", prio="高",
        case="BS-IT-044",
        symptom="搜索含空格或 & 的关键词后点击\"下一页\"，搜索条件被截断，翻页结果变成无条件列表或错误结果。",
        steps="1. 访问 http://localhost:8080/Book/client/bookServlet?action=pageByNameOrAuthor&nameorauthor=%E4%B8%89%20%E4%BD%93（搜索\"三 体\"）\n2. 查看页面翻页链接：href 形如 client/bookServlet?action=pageByNameOrAuthor&nameorauthor=三 体&pageNo=2\n3. 点击\"下一页\"，观察搜索结果",
        expect="翻页链接对参数值做 URL 编码（空格→%20，&→%26），点击后 nameorauthor=三 体 完整传递。",
        actual="href 中搜索词以原始字符拼接，空格未编码导致点击后参数被截断为\"三\"；含 & 的搜索词（如 a&b）后段被解析为独立参数，搜索条件丢失。",
        root="src/com/yj/web/ClientBookServlet.java 第 46-49 行：sb.append(\"&nameorauthor=\").append(req.getParameter(\"nameorauthor\")) 直接拼接原始值，未调用 URLEncoder.encode。",
        fix="拼接前对参数值 URLEncoder.encode(v, \"UTF-8\")；pageByPrice 的 min/max 拼接同理。",
        verify="修复后搜索\"三 体\"点击下一页，地址栏参数完整、结果仍按关键词过滤；run-tests.bat 复测 test044 应通过。"),
    dict(
        id="BUG-M3-03", title="价格区间 min>max 无校验、无提示，静默返回空列表",
        module="M3 图书浏览与检索", sev="中", prio="中",
        case="BS-IT-045",
        symptom="设置 min=100&max=10 后页面显示空列表，无任何错误或提示信息，用户无从得知区间设置错误。",
        steps="1. 访问 http://localhost:8080/Book/client/bookServlet?action=pageByPrice&min=100&max=10\n2. 观察页面：空列表、无提示",
        expect="给出可见提示（如\"价格区间设置有误\"）或自动交换 min/max 归一化处理。",
        actual="BETWEEN 100 AND 10 恒为假，查询结果为空，页面静默显示无图书。",
        root="src/com/yj/web/ClientBookServlet.java pageByPrice() 与 BookServiceImpl.pageByPrice()：对 min>max 未做任何校验。",
        fix="Service 层加 if (min > max) 交换两值或向前台返回提示信息。",
        verify="修复后 min=100&max=10 应显示提示或返回 10~100 区间图书；run-tests.bat 复测 test045 相应断言应通过。"),
    dict(
        id="BUG-M3-04", title="价格筛选参数仅支持整数，小数价格静默失效",
        module="M3 图书浏览与检索", sev="中", prio="中",
        case="BS-IT-045",
        symptom="筛选 min=56.5&max=56.5（库中《三体》恰为 56.5 元）时，返回的不是该书，而是全部图书的第一页——筛选条件被完全忽略。",
        steps="1. 访问 http://localhost:8080/Book/client/bookServlet?action=pageByPrice&min=56.5&max=56.5\n2. 统计页面图书数：显示 4 本（默认每页全部图书）",
        expect="价格是 DECIMAL 类型，小数参数应正常参与筛选，精确命中 56.5 元的书。",
        actual="WebUtils.parseInt 按 int 解析 \"56.5\" 抛 NumberFormatException，被 catch 后静默返回默认值 min=0、max=Integer.MAX_VALUE，等价于取消筛选。",
        root="src/com/yj/utils/WebUtils.java parseInt() 仅支持整数；ClientBookServlet.pageByPrice() 用它解析本来就是小数的价格参数。",
        fix="价格参数改用 new BigDecimal(str) 解析（异常时回退默认 0~MAX）。",
        verify="修复后 min=56.5&max=56.5 仅返回《三体》；run-tests.bat 复测 test045 应通过。"),
    dict(
        id="BUG-M3-05", title="销量榜单 SQL 使用 LIMIT 1,50，跳过销量第一名",
        module="M3 图书浏览与检索", sev="中", prio="中",
        case="BS-IT-046",
        symptom="销量排行榜不显示销量最高的图书，榜单从真正的第二名开始。",
        steps="1. 将某书（如 id=1《解忧杂货店》）销量改为全库最高：UPDATE t_book SET sales=999 WHERE id=1\n2. 访问 http://localhost:8080/Book/client/bookServlet?action=pageOrder\n3. 榜单中找不到《解忧杂货店》；从原第二名开始展示\n4. 测试后还原 UPDATE t_book SET sales=100 WHERE id=1",
        expect="榜单包含销量第一名，并按销量降序排列。",
        actual="SQL 为 SELECT * FROM t_book ORDER BY sales DESC LIMIT 1,50，偏移量 1 跳过第一行。",
        root="src/com/yj/dao/impl/BookDaoImpl.java 第 88 行 queryForPageItemsOrder()。",
        fix="LIMIT 1,50 改为 LIMIT 50。",
        verify="修复后重复上述步骤，榜单第一名即销量最高的书；run-tests.bat 复测 test046 应通过。"),
    dict(
        id="BUG-M6-01", title="新增图书无服务端输入校验：空书名/负价格/非法价格直接入库",
        module="M6 后台管理", sev="高", prio="高",
        case="BS-IT-102",
        symptom="后台新增图书时提交空书名、负价格或非数字价格，系统不提示错误并直接写库：空书名存为空串、负价存为 -5.00、非数字价格存为 NULL。",
        steps="1. admin 登录后台，进入图书管理→新增（或直接 POST /Book/manager/bookServlet?action=add）\n2. 书名留空，其他正常，提交 → 新增成功\n3. 查库 SELECT id,name,price FROM t_book ORDER BY id DESC：出现 name='' 的行\n4. 再分别提交 price=-5 与 price=abc：-5.00 入库；abc 行的 price 为 NULL",
        expect="服务端校验必填与数值合法性：空书名、负价格、非数字价格应拒绝并提示。",
        actual="WebUtils.copyParamToBean 中 BeanUtils 转换异常仅 printStackTrace 被吞，其余值原样写库。",
        root="src/com/yj/web/BookServlet.add() 无校验；src/com/yj/utils/WebUtils.java copyParamToBean() 吞异常。",
        fix="add/update 入口对 name 非空、price≥0 且为合法数字、sales/stock≥0 做校验，不合法带错误信息转发回编辑页。",
        verify="修复后三类非法输入均被拒绝且有提示；run-tests.bat 复测 test102 应通过。"),
    dict(
        id="BUG-M6-02", title="删除图书：非数字 id 引发 500；不存在的 id 静默无提示",
        module="M6 后台管理", sev="中", prio="中",
        case="BS-IT-104",
        symptom="删除图书传非数字 id 时服务端 500；传不存在的大 id 时静默跳回列表，用户不知道删除并未生效。",
        steps="1. admin 登录后访问 /Book/manager/bookServlet?action=delete&id=abc → HTTP 500（NumberFormatException）\n2. 访问 /Book/manager/bookServlet?action=delete&id=999999 → 正常跳回列表，无任何提示",
        expect="非数字 id 提示参数错误；不存在的 id 提示\"图书不存在/删除失败\"。",
        actual="Integer.parseInt(\"abc\") 抛异常未捕获直达 500；delete 影响 0 行时无差异处理。",
        root="src/com/yj/web/BookServlet.delete() 第 34-35 行：Integer.parseInt 未防护、返回值未检查。",
        fix="用 WebUtils.parseInt 容错解析；根据 deleteBookById 返回的影响行数给出相应提示。",
        verify="修复后 id=abc 与 id=999999 均有友好提示且无 500；run-tests.bat 复测 test104 应通过。"),
    dict(
        id="BUG-M6-03", title="新增重复用户名触发唯一约束，服务端 500 错误页",
        module="M6 后台管理", sev="高", prio="高",
        case="BS-IT-105",
        symptom="后台用户管理新增已存在的用户名（如 admin）时，返回 Tomcat 500 错误页（堆栈暴露），无友好提示。",
        steps="1. admin 登录后 POST /Book/manager/UserServlet?action=add，参数 username=admin&password=x&email=dup@test.com\n2. 观察响应：HTTP 500，页面为服务器错误页\n3. Tomcat 日志出现 SQLIntegrityConstraintViolationException: Duplicate entry 'admin' for key 't_user.username'",
        expect="提示\"用户名已存在\"并停留在新增页（与前台注册的 AJAX 查重体验一致）。",
        actual="UserDaoImpl.addUser 抛出的唯一约束异常经 TransactionFilter 包装后直接抛给容器，回滚事务并显示 500 错误页。",
        root="src/com/yj/web/ManagerUserServlet.add() 调用链无重复检查、无异常处理。",
        fix="add 前先 userService.existsUsername 判断；或捕获约束冲突异常转发带提示的表单页。",
        verify="修复后重复用户名返回带提示的页面（非 500）；正常新增/删除不受影响；run-tests.bat 复测 test105 应通过。"),
]


def add_kv(doc, k, v):
    p = doc.add_paragraph()
    r = p.add_run(k + "：")
    r.bold = True
    r.font.size = Pt(10.5)
    p.add_run(v).font.size = Pt(10.5)


def main():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(10.5)

    h = doc.add_heading("网上书店系统 缺陷报告（丙：M3 图书浏览与检索 + M6 后台管理）", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.add_run("被测系统：网上书店（JavaWeb MVC，Tomcat 9 + MySQL 8）    测试日期：2026-09-14/15    用例范围：BS-IT-041~046、BS-IT-101~106\n"
              "缺陷总数：8（高 4 / 中 4），全部经自动化用例实测确认，复现步骤可直接执行。"
              "每条含现象、复现步骤、根因、修复建议与修复验证方式；修复后以 tests\\run-tests.bat 复测对应用例应为通过。").font.size = Pt(9)

    doc.add_heading("缺陷汇总", level=1)
    table = doc.add_table(rows=1, cols=6)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, t in enumerate(["编号", "标题", "模块", "严重度", "优先级", "关联用例"]):
        hdr[i].text = t
    for d in DEFECTS:
        cells = table.add_row().cells
        cells[0].text = d["id"]
        cells[1].text = d["title"]
        cells[2].text = d["module"]
        cells[3].text = d["sev"]
        cells[4].text = d["prio"]
        cells[5].text = d["case"]

    for d in DEFECTS:
        doc.add_heading(f"{d['id']}  {d['title']}", level=1)
        add_kv(doc, "所属模块", f"{d['module']}（{d['case']}）")
        add_kv(doc, "严重程度 / 优先级", f"{d['sev']} / {d['prio']}")
        add_kv(doc, "现象", d["symptom"])
        add_kv(doc, "复现步骤", d["steps"])
        add_kv(doc, "预期结果", d["expect"])
        add_kv(doc, "实际结果与根因", d["actual"])
        add_kv(doc, "代码位置", d["root"])
        add_kv(doc, "修复建议", d["fix"])
        add_kv(doc, "修复验证", d["verify"])

    doc.save(OUT)
    print(f"saved: {OUT}, defects={len(DEFECTS)}")


if __name__ == "__main__":
    main()
