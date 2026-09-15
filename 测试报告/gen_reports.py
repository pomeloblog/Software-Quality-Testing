# -*- coding: utf-8 -*-
"""
按附录3（模块一测试报告）与附录4（模块二AI融合报告）模板生成全组测试报告。

数据来源：
  - 甲（王尚清）M1/M4：origin/wsq 分支 测试用例清单-M1-M4(甲).xlsx / 测试报告 / 缺陷报告
  - 丙 M3/M6：本仓库 测试用例/丙-* （12条用例实测：5 OK / 7 NG，8缺陷）
  - 乙 M2/M5：占位待补
用法：python gen_reports.py
"""
import copy
import os

from docx import Document
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
TPL_DIR = r"E:\软件质量测试\软件测试实践作业\实践作业-文档模板2026"
OUT1 = os.path.join(HERE, "测试报告（模块一）.docx")
OUT2 = os.path.join(HERE, "测试报告（模块二-AI融合实践）.docx")

# ============================================================
# 通用工具
# ============================================================

def is_toc_para(p):
    s = p.text
    return "\t" in s and s.strip()[-1:].isdigit()


def body_paras(doc):
    return [p for p in doc.paragraphs if not is_toc_para(p)]


def replace_body_text(doc, prefix, new_text, Occur=1):
    """把正文中以 prefix 开头的段落的文本替换为 new_text（保留首个 run 的格式）。"""
    n = 0
    for p in body_paras(doc):
        if p.text.strip().startswith(prefix):
            if not p.runs:
                p.add_run(new_text)
            else:
                p.runs[0].text = new_text
                for r in p.runs[1:]:
                    r.text = ""
            n += 1
            if n >= Occur:
                return p
    return None


def insert_table_after(anchor_para, doc, data, style="Table Grid", font_size=9):
    """在 anchor 段落后插入表格（doc.add_table 追加到文末后移动元素）。"""
    t = doc.add_table(rows=len(data), cols=len(data[0]))
    try:
        t.style = style
    except Exception:
        pass
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            cell = t.cell(i, j)
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(v))
            run.font.size = Pt(font_size)
    anchor_para._p.addnext(t._tbl)
    return t


def fill_cover(doc, system, extra_rev=()):
    for p in doc.paragraphs:
        s = p.text.strip()
        if s == "XX系统":
            if p.runs:
                p.runs[0].text = system
                for r in p.runs[1:]:
                    r.text = ""
    rev = doc.tables[0]
    rows = [("2026-09-15", "V1.0", "初稿：丙负责的 M3 图书浏览与检索 + M6 后台管理（12条用例实测）", "丙")]
    rows += [("2026-09-15", "V1.1", "8个缺陷全部修复（v1.1-fix）并回归12/12通过，质量结论更新为达标", "丙")]
    rows += list(extra_rev)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            rev.cell(i + 1, j).text = str(v)


def update_toc(path):
    import win32com.client
    app = win32com.client.Dispatch("KWPS.Application")
    app.Visible = False
    d = app.Documents.Open(path)
    try:
        d.Fields.Update()
        for i in range(1, d.TablesOfContents.Count + 1):
            d.TablesOfContents(i).Update()
    except Exception as e:
        print("TOC update warn:", e)
    d.Save()
    d.Close(False)
    app.Quit()


# ============================================================
# 报告一：附录3 模块一测试报告
# ============================================================

OVERVIEW = (
    "被测系统为网上书店系统（Java/Servlet/JSP，MVC 三层架构：bean→dao→service→web，无构建工具，"
    "依赖 jar 位于 web/WEB-INF/lib）。系统分为前台（浏览/检索/购物车/下单）与后台（图书/用户/订单管理）"
    "两大部分，服务器 Tomcat 9（context path=/Book），数据库 MySQL（库名 book，4 张表：t_user、t_book、"
    "t_order、t_order_item），源码共 46 个 Java 文件。本报告为本组成员丙负责的两个关键模块："
)

MODULE_TBL = [
    ["模块", "核心功能", "主要源文件", "代码规模", "负责人"],
    ["M3 图书浏览与检索", "分页/搜索/价格筛选/销量榜单", "ClientBookServlet、BookServiceImpl、BookDaoImpl", "约450行", "丙"],
    ["M6 后台管理", "图书/用户增删改查/发货/总账单", "BookServlet、ManagerUserServlet、ManagerOrderServlet", "约500行", "丙"],
]

SCOPE_TBL = [
    ["项目", "说明"],
    ["本次测试覆盖",
     "M3 图书浏览与检索：首页分页（pageNo/pageSize 边界）、按书名/作者搜索（等价类与特殊字符）、"
     "翻页链接条件保持、价格区间筛选、销量榜单排序与完整性；"
     "M6 后台管理：管理员登录与图书管理列表、新增图书输入校验、修改图书回显与更新、删除图书边界、"
     "用户管理新增与重复用户名、总账单对账与发货状态流转。共 2 个模块 12 条用例，全部自动化。"],
    ["本次测试不覆盖",
     "① M1 用户认证、M2 权限控制、M4 购物车、M5 订单与事务——由小组其他成员负责，不在本报告范围；"
     "② 后台管理的图书列表多页翻页组合、订单详情页样式——优先级较低，用例额度分配给了缺陷更集中的路径；"
     "③ 性能、并发、兼容性、前端样式——超出本阶段（模块一：测试基础实践）范畴。"],
    ["不覆盖原因", "分工安排 / 优先级排序 / 阶段范围外 / 环境限制。"],
]

STRATEGY = (
    "黑盒方法按被测对象的输入域特征选择，本报告使用等价类划分、边界值分析、场景法三种（满足≥2种要求）：\n"
    "（1）等价类划分：用于输入域可明显划分合法/非法类的功能，如搜索词（命中书名/命中作者/不存在词/注入串）、"
    "新增图书（合法值/空值/负值/类型非法值）、新增用户（新用户名/重复用户名）——从每类取代表值暴露校验缺失。\n"
    "（2）边界值分析：用于有数值边界的参数，如分页 pageNo/pageSize（0、-1、超大、非数字）、"
    "价格区间 min/max（相等、交叉、小数）、删除图书 id（存在/极大不存在/非数字）——7 条 NG 用例中 5 条由边界输入触发，"
    "验证了边界值法对本系统的针对性。\n"
    "（3）场景法：用于多步骤业务流程，如管理员登录→后台管理全流程（含未登录对照组）、搜索→翻页链接完整性、"
    "总账单对账（页面统计 vs 直查库）、发货状态流转（status 0→1）。\n"
    "白盒手段上，测试前对被测模块源码走查（分页取模运算、URL 拼接、SQL 常量），指导用例设计（如 LIMIT 1,50 可疑点）；"
    "执行层为 HTTP 集成测试：HttpURLConnection 向运行中的 /Book 发真实请求，并以 JDBC 直连数据库断言落库结果。"
)

ENV_TEXT = (
    "被测系统与测试代码运行条件（本机）：操作系统 Windows 11 家庭版；JDK 11（测试进程，编译 -encoding UTF-8）；"
    "Tomcat 9（context path=/Book，端口 8080）；MySQL 8.0（库名 book）。"
    "因项目自带 mysql-connector 5.1.7 驱动不支持 MySQL 8.0 默认的 caching_sha2_password 认证，"
    "创建专用账号 bookstore/123456（mysql_native_password，仅授权 book 库）适配，同步修改 src/jdbc.properties。\n"
    "测试代码运行条件：JUnit 4.12 + hamcrest-core 1.3 + mysql-connector 5.1.7（已复制到 tests/lib/，无需 Maven）；"
    "一键运行脚本 tests/run-tests.bat（编译 src/*.java → JUnitCore 执行 → 输出逐用例结果与汇总）。\n"
    "被测应用账号：管理员 admin/admin；测试造数：cp_test_01 用户、CPTEST/CPDEL 标记图书、CPTEST_ORDER_1 临时订单（全部用后清理）。"
    "数据库由 sql/bookstore.sql 初始化（29 本图书、admin 用户、sales/stock 各 100）。"
)

CASE_OVERVIEW = [
    ["测试方法", "用例数量", "占比", "已自动化数"],
    ["等价类划分", "3", "25.0%", "3"],
    ["边界值分析", "4", "33.3%", "4"],
    ["场景法", "5", "41.7%", "5"],
    ["合计", "12", "100%", "12（100%自动化）"],
]

TYPICAL_CASES = [
    ("BS-IT-041 首页分页 pageNo 边界容错（边界值法）",
     "输入：pageNo=1/0/-1/99999/abc。预期：越界与非法值分别回退第 1 页/末页，非数字取默认 1，全部 HTTP 200。"
     "实际：与预期一致（本模块为数不多的容错正确点，作为对照）。"
     "设计依据：pageNo 经 WebUtils.parseInt 解析且 Service 层有 pageNo<1 / >pageTotal 双向归一，边界值五点全覆盖验证该归一逻辑。"),
    ("BS-IT-042 分页 pageSize 边界（边界值法）",
     "输入：pageSize=6/abc/0/-1。预期：非法值容错为默认每页 4 条，不应服务器错误。"
     "实际：pageSize=0 返回 500（pageTotalCount%pageSize 除零 ArithmeticException）；-1 返回 500（LIMIT 负数 SQL 错误）。"
     "设计依据：BookServiceImpl.page() 中 pageSize 直接参与取模与 LIMIT，为典型边界缺陷点。"),
    ("BS-IT-044 搜索翻页链接未 URL 编码（场景法/缺陷定向）",
     "输入：nameorauthor=\"三 体\"（含空格）与 a&b（含连接符）搜索后检查翻页 href。"
     "预期：参数值 URL 编码，翻页后搜索条件完整。"
     "实际：href 中原始字符拼接（ClientBookServlet 第 48 行），点击下一页后搜索词截断、条件丢失。"
     "设计依据：代码走查发现分页 URL 拼接未编码，属缺陷定向验证。"),
    ("BS-IT-045 价格区间筛选边界（边界值+等价类）",
     "输入：min,max=(10,30)/(57,57)/(100,10)/(56.5,56.5)。预期：区间结果与库一致、精确匹配、交叉区间有提示、小数正常筛选。"
     "实际：前两项正确；min>max 静默空列表无提示；56.5 被 parseInt 按整数解析失败后静默回退为无筛选（页面显示全部图书）。"
     "设计依据：价格为 DECIMAL 而解析按 int，类型不匹配是天然边界缺陷点。"),
    ("BS-IT-046 销量榜单跳过第一名（缺陷定向）",
     "输入：造数将 id=1 图书销量改为全库最高（999），请求榜单后还原。"
     "预期：榜单含销量第一名。实际：不含——SQL 为 ORDER BY sales DESC LIMIT 1,50，偏移量 1 跳过第一名"
     "（BookDaoImpl 第 88 行）。设计依据：代码走查发现 LIMIT 常量可疑，通过造数使缺陷可见。"),
    ("BS-IT-102 新增图书输入校验（等价类法）",
     "输入：合法图书 / 空书名 / 负价格 / 非数字价格四类代表值。"
     "预期：后三类被服务端校验拒绝。实际：全部直接入库（空名存空串、负价入库、abc 以 NULL 入库）。"
     "设计依据：add() 链路无任何校验，BeanUtils 转换异常被吞，属校验缺失型缺陷。"),
]

DIST_TBL = [
    ["序号", "被测模块", "需求简述", "测试用例数量"],
    ["1", "M3 图书浏览与检索", "分页参数容错、搜索模糊匹配与注入防护、翻页条件保持、价格区间筛选、榜单正确性", "6"],
    ["2", "M6 后台管理", "管理员访问控制（对照）、图书增删改查合法性、编辑回显、重复用户名处理、发货流转、账单对账", "6"],
    ["合计", "—", "—", "12"],
]

FRAMEWORK = (
    "测试框架：JUnit 4.12（项目自带 jar，不引入 Maven，符合无构建工具的 Java EE 工程现状）。\n"
    "工程结构（tests/ 目录）：\n"
    "├── run-tests.bat          一键运行：编译 → JUnitCore 执行 → 输出逐用例结果与汇总\n"
    "├── lib/                   junit-4.12 / hamcrest-core-1.3 / mysql-connector-5.1.7（自包含，克隆即用）\n"
    "└── src/\n"
    "    ├── HttpTestUtil.java      测试基础设施：HTTP 请求（会话保持+手动重定向）+ JDBC 造数断言\n"
    "    ├── M3BookBrowseTest.java  M3 用例 BS-IT-041~046\n"
    "    └── M6ManagerTest.java     M6 用例 BS-IT-101~106（@Before 自动登录 admin）\n"
    "测试路线为黑盒集成测试：HttpURLConnection 向运行中的 /Book 发真实 HTTP 请求（GET/POST、带会话、带 Referer），"
    "以 JDBC 直连数据库断言落库结果；断言口径统一为\"系统应有的正确行为\"，断言失败即 NG 用例（缺陷证据），"
    "JUnit 输出与 Excel 用例清单 Result 列一一对应。"
)

KEY_IMPL = (
    "核心设计一：HTTP 会话保持与重定向处理（丙，HttpTestUtil，节选）：\n"
    "  // 手动跟随302：JDK自动跟随不传Cookie，登录页新发的JSESSIONID会污染会话（已踩坑）\n"
    "  for (int hop = 0; hop < 5; hop++) {\n"
    "      conn = (HttpURLConnection) new URL(url).openConnection();\n"
    "      conn.setInstanceFollowRedirects(false);\n"
    "      if (jsessionId != null) conn.setRequestProperty(\"Cookie\", \"JSESSIONID=\" + jsessionId);\n"
    "      int code = conn.getResponseCode();\n"
    "      String location = conn.getHeaderField(\"Location\");\n"
    "      if ((code == 302) && location != null) {\n"
    "          url = location.startsWith(\"http\") ? location\n"
    "              : location.startsWith(\"/\") ? SERVER_ROOT + location : BASE + \"/\" + location;\n"
    "          curMethod = \"GET\"; curBody = null; continue;   // 302后按浏览器惯例降级GET\n"
    "      }\n"
    "      return new Response(code, readAll(...));\n"
    "  }\n"
    "设计要点：(1)全部用例参数化集中在 URL 拼接层，用例方法只给输入与断言；"
    "(2)数据库断言与造数统一走 HttpTestUtil.exec/queryLong（预编译参数）；"
    "(3)改库用例以 try/finally + @AfterClass 双保险还原（如销量改999后还原、测试图书按作者标记清理），保证可重复执行；"
    "(4)外部依赖不 Mock 数据库——被测系统与断言共用同一真实库，验证真实落库；验证码等不可控依赖才用 Stub 替代。"
)

RESULT_TEXT = (
    "分模块执行结果如下（初测=缺陷发现阶段；修复后回归=8 个缺陷修复并重新运行全部用例）："
)

RESULT_TBL = [
    ["模块", "用例数", "初测通过", "初测失败", "初测通过率", "修复后回归"],
    ["M3 图书浏览与检索", "6", "2", "4", "33.3%", "6/6 全通过"],
    ["M6 后台管理", "6", "3", "3", "50.0%", "6/6 全通过"],
    ["合计", "12", "5", "7", "41.7%", "12/12 全通过（100%）"],
]

DEFECT_TBL = [
    ["编号", "缺陷简述", "模块", "严重度", "发现方式", "状态"],
    ["BUG-M3-01", "pageSize=0/-1 服务端500（除零/SQL错误）", "M3", "高", "自动化", "已修复关闭"],
    ["BUG-M3-02", "搜索翻页链接未URL编码，条件丢失", "M3", "高", "自动化", "已修复关闭"],
    ["BUG-M3-03", "价格区间 min>max 无校验无提示", "M3", "中", "自动化", "已修复关闭"],
    ["BUG-M3-04", "价格参数仅支持整数，小数静默失效", "M3", "中", "自动化", "已修复关闭"],
    ["BUG-M3-05", "销量榜单 LIMIT 1,50 跳过第一名", "M3", "中", "自动化", "已修复关闭"],
    ["BUG-M6-01", "新增图书无校验：空名/负价/NULL价入库", "M6", "高", "自动化", "已修复关闭"],
    ["BUG-M6-02", "删除图书非数字id 500；不存在id静默", "M6", "中", "自动化", "已修复关闭"],
    ["BUG-M6-03", "新增重复用户名唯一约束冲突 500", "M6", "高", "自动化", "已修复关闭"],
]

ROOT_CAUSE = (
    "分析一（BUG-M3-01，pageSize=0/-1 服务端 500，高/健壮性）——完整闭环："
    "①发现：BS-IT-042 以 pageSize=0 请求分页返回 HTTP 500；"
    "②定位：BookServiceImpl.page() 第 59 行 pageTotalCount % pageSize 抛除零 ArithmeticException；"
    "pageSize=-1 时 pageTotal 为负、pageNo 归一为负，LIMIT 起始/条量为负触发 MySQL 语法错误；"
    "③根因：parseInt 已容错非数字，但解析结果未做下界校验；"
    "④修复：page()/pageByPrice()/pageByNameOrAuthor() 三个方法入口统一增加 if (pageSize < 1) pageSize = Page.PAGE_SIZE;；"
    "⑤回归验证：run-tests.bat 复测 test042 通过——pageSize=0/-1/abc 均返回 200 且每页 4 条。\n"
    "分析二（BUG-M6-01，新增图书无输入校验，高/数据正确性）——完整闭环："
    "①发现：BS-IT-102 空书名/负价格/非数字价格三类输入全部直接入库（负价 -5.00、abc 以 NULL 入库）；"
    "②定位：BookServlet.add() 无校验；WebUtils.copyParamToBean() 中 BeanUtils 转换异常仅 printStackTrace 被吞；"
    "③根因：全链路无服务端校验层，异常被静默吞掉后脏数据直接落库；"
    "④修复：新增 validateBook()（书名非空、价格≥0 且合法数字、销量/库存≥0），add/update 入库前调用，"
    "不合法转发回 book_edit.jsp 显示红色提示，并在 JSP 增加提示区；"
    "⑤回归验证：test102 通过——三类非法输入均被拒绝且库中无脏数据，合法图书正常入库。\n"
    "分析三（BUG-M3-05，榜单跳过第一名，中/功能正确性）——完整闭环："
    "①发现：BS-IT-046 造数（销量改为全库最高）后榜单不含第一名；"
    "②定位：BookDaoImpl.queryForPageItemsOrder() 第 88 行 SQL 为 ORDER BY sales DESC LIMIT 1,50；"
    "③根因：LIMIT 偏移量 1 恰好跳过销量最高的一行；"
    "④修复：SQL 改为 LIMIT 50（一行改动）；"
    "⑤回归验证：test046 通过——造数后销量第一名《解忧杂货店》出现在榜单首位。"
)

QUALITY = (
    "M3 图书浏览与检索、M6 后台管理共 12 条用例，初测通过 5 条（41.7%），确认有效缺陷 8 个（高 4 / 中 4）。\n"
    "从缺陷分布看，正常功能路径（分页浏览、搜索命中、合法数据增删改、账单统计、发货流转）实现正确——"
    "总账单四项统计与直查数据库完全一致、发货状态流转正常；初测暴露的问题集中在异常输入防御："
    "8 个缺陷中 6 个属于\"参数校验缺失\"（pageSize 下界、空书名、负价、min>max、非数字 id、重复用户名），"
    "2 个属于实现细节错误（URL 未编码、LIMIT 偏移）。\n"
    "修复与回归：8 个缺陷已于 2026-09-15 全部修复（build v1.1-fix），重新运行 tests\\run-tests.bat，"
    "12 条用例全部通过（通过率 100%），无回归缺陷，测试数据库状态完全还原。\n"
    "结论：修复后 M3、M6 两个模块功能与健壮性均达到发布标准，全部缺陷关闭。"
)

CONTRIB = [
    ["成员", "承担模块", "具体任务", "测试代码行数", "发现缺陷数", "工作量占比"],
    ["丙", "M3 图书浏览与检索\nM6 后台管理", "测试环境搭建（MySQL8.0认证适配、Tomcat部署、Git仓库与分支管理）；12条用例设计与自动化（HttpTestUtil+run-tests.bat一键运行）；8个缺陷确认并定位至代码行；用例清单Excel、缺陷报告（附录2模板）、本测试报告", "约700行", "8", "100%（本报告范围）"],
]


def build_report1():
    doc = Document(os.path.join(TPL_DIR, "附录3：测试报告模板（模块一）.docx"))
    fill_cover(doc, "网上书店系统")
    # 先捕获模板表格引用（后续插入新表会使 doc.tables 索引偏移）
    scope_t, overview_t, dist_t = doc.tables[1], doc.tables[2], doc.tables[3]

    # 1 概述：替换占位段 + 模块表
    p = replace_body_text(doc, "说明被测软件系统名称", OVERVIEW)
    insert_table_after(p, doc, MODULE_TBL)

    # 2.1 范围表
    for i, row in enumerate(SCOPE_TBL):
        if i == 0:
            continue
        for j, v in enumerate(row):
            scope_t.cell(i, j).text = v

    # 2.2 策略
    replace_body_text(doc, "在此说明针对不同测试项", STRATEGY)

    # 2.3 环境
    p = replace_body_text(doc, "在此说明程序的运行条件", ENV_TEXT)
    # 删除第二个占位段（"并说明测试代码的运行条件"）已并入上文
    replace_body_text(doc, "并说明测试代码的运行条件", "")

    # 3.1 用例总览表（需加行）
    t = overview_t
    for i, row in enumerate(CASE_OVERVIEW):
        if i == 0:
            continue
        if i == 1:
            for j, v in enumerate(row):
                t.cell(1, j).text = v
        else:
            cells = t.add_row().cells
            for j, v in enumerate(row):
                cells[j].text = v

    # 3.2 典型用例
    txt = "\n".join(f"{i+1}. {t_}" for i, t_ in enumerate(TYPICAL_CASES))
    replace_body_text(doc, "针对不同被测模块，共选取", txt)

    # 3.3 用例分布表（需加行）
    t = dist_t
    for i, row in enumerate(DIST_TBL):
        if i == 0:
            continue
        if i <= 2:
            for j, v in enumerate(row):
                t.cell(i, j).text = v
        else:
            cells = t.add_row().cells
            for j, v in enumerate(row):
                cells[j].text = v

    # 4.1 / 4.2
    replace_body_text(doc, "测试框架：如 pytest", FRAMEWORK)
    replace_body_text(doc, "选取核心的测试代码片段", KEY_IMPL)

    # 5.1 执行结果（文字+表格）
    p = replace_body_text(doc, "分不同的功能模块，说明总用例数", RESULT_TEXT)
    insert_table_after(p, doc, RESULT_TBL)

    # 5.2 缺陷摘要
    p = replace_body_text(doc, "针对每个缺陷，简要描述", "初测共确认有效缺陷 8 个（高 4 / 中 4），全部由自动化用例执行发现；已于 2026-09-15 全部修复并回归验证后关闭。摘要如下：")
    insert_table_after(p, doc, DEFECT_TBL)

    # 5.3 根因分析
    replace_body_text(doc, "选取2-3个最有代表性的缺陷", ROOT_CAUSE)

    # 5.4 质量结论
    replace_body_text(doc, "基于测试结果，得出被测模块", QUALITY)

    # 6 贡献度：模板只有占位段，插入表格
    p = replace_body_text(doc, "说明每个成员承担的具体任务", "本报告覆盖范围的全部工作由丙独立完成：")
    insert_table_after(p, doc, CONTRIB)

    doc.save(OUT1)
    print("saved:", OUT1)


# ============================================================
# 报告二：附录4 模块二 AI 融合报告（方案2：测试手段含AI）
# ============================================================

AI_OVERVIEW = (
    "被测系统为网上书店系统（JavaWeb MVC，46 个 Java 文件，Tomcat 9 + MySQL），小组选择 6 个模块中的 4 个"
    "（M1/M3/M4/M6，另 M2/M5 由乙人工完成）开展测试。本模块二实践采用方案2（测试手段含 AI）："
    "在传统测试流程中系统性引入 AI 编程助手辅助完成环境搭建、测试用例设计、自动化脚本开发、缺陷定位与文档编写，"
    "并对 AI 产出进行人工复核与修正。选择方案 2 的理由：被测对象为传统确定性 JavaWeb 系统（无 AI 模块），"
    "而测试过程中存在大量重复性机械劳动（环境适配、脚本编写、文档整理），是 AI 辅助增效的典型场景；"
    "同时小组成员可完整观察\"AI 产出—人工复核—修正\"的协作过程，为反思提供真实素材。"
    "使用工具：Claude Code（命令行 AI 编程助手，Anthropic），全程对话记录留存于会话日志。"
)

AI_PROCESS = (
    "实践按\"AI 执行→人工复核→修正定稿\"循环推进，关键过程如下：\n"
    "（1）环境搭建：AI 完成 Git 仓库初始化与 GitHub 连接、MySQL 8.0 认证适配（识别出项目自带 mysql-connector 5.1.7 "
    "与 8.0 默认认证不兼容，创建 mysql_native_password 专用账号）、编译部署到 Tomcat 并验证登录流程。"
    "人工复核：确认数据库导入结果与页面数据一致。【此处插入截图1：环境搭建对话，待补】\n"
    "（2）用例设计：AI 依据分工文档的模块定义与代码走查结果，生成 M3/M6 共 12 条用例（边界值 4、等价类 3、场景法 5），"
    "编号 BS-IT-041~046/101~106，按模板字段输出 Excel。人工修正：用例取舍与预期结果的产品语义判断（如\"删除不存在的 id "
    "是否必须提示\"由人拍板）。【此处插入截图2：用例生成对话，待补】\n"
    "（3）脚本开发：AI 实现 HttpTestUtil/M3/M6 测试类与 run-tests.bat 一键运行；实测 12 条得 5 OK / 7 NG。\n"
    "（4）缺陷定位：AI 将 7 条 NG 用例全部定位到具体代码行（如 BookServiceImpl 第 59 行除零、"
    "BookDaoImpl 第 88 行 LIMIT 1,50），并按附录2 模板生成 8 个缺陷的缺陷报告。"
    "人工复核：对关键缺陷用 curl 独立复现确认。【此处插入截图3：缺陷定位对话，待补】\n"
    "（5）文档统稿：AI 按附录2/3/4 模板生成缺陷报告与测试报告并更新目录域。"
)

AI_RESULTS = (
    "新增测试成果：M3/M6 共 12 条自动化用例（100% 自动化）与 8 个实测确认缺陷（高 4 / 中 4），"
    "全部有可复现步骤与代码级根因，详见 测试用例/丙-测试用例清单-M3-M6.xlsx 与 丙-缺陷报告-M3-M6.docx。\n"
    "AI 辅助与人工改进的关键差异对比：\n"
    "①效率：环境搭建+脚本开发+三份文档，AI 辅助下约半天完成，估计纯人工需 2-3 天；\n"
    "②深度：AI 主动走查源码发现的缺陷（榜单 LIMIT 1、翻页未编码）超出用例清单原始线索；\n"
    "③准确性：AI 首版测试工具存在会话保持缺陷导致 1 条用例假通过（详见第 3 节），经人工发现异常数据后复查修正——"
    "修正后该用例（BS-IT-102）才暴露出真实缺陷，说明 AI 产出的测试代码本身也必须被验证；\n"
    "④判断力：严重程度定级、用例取舍、模板符合性等需要人对课程要求的理解，AI 仅能给出建议。"
)

AI_REFLECT = (
    "AI 工具帮我们做了什么：环境适配（MySQL 8 认证兼容方案）、46 个源文件的快速走查、12 条用例设计与三份模板文档生成、"
    "8 个缺陷的根因定位（精确到代码行）与修复建议、一键测试脚本（run-tests.bat）。\n"
    "AI 工具哪里\"不靠谱\"（真实案例）：①HttpTestUtil 首版使用 JDK 自动跟随 302 重定向，不传 Cookie，"
    "被登录页新发的 JSESSIONID 污染会话，导致 BS-IT-102（新增图书无校验）连续两轮假通过——AI 自己的测试代码引入了新的缺陷模式，"
    "直到人工发现\"用例通过但数据库无新增记录\"的矛盾才推动排查；②生成缺陷报告时按标题前缀查找正文段落，"
    "误匹配到目录条目，将\"1 引言\"标题覆盖为结果分析内容（已修正为过滤目录段）；③Excel 首版列数与模板不符（14 列 vs 13 列）。"
    "共性：AI 错误多发生在\"对既有约定（模板结构、HTTP 语义）的隐性假设\"上，且错误输出表面通常\"看起来很对\"。\n"
    "人做了什么：发现用例结果与数据库状态的矛盾并要求复查（假通过的唯一线索来自人的交叉核对）；"
    "用 curl 独立复现关键缺陷；决定用例预期结果的产品语义（提示 vs 静默、定级高/中）；核对模板符合性与提交规范。\n"
    "AI 工具在测试中的角色定位：辅助者。判断依据：AI 在机械性、模式化劳动上（编译部署、脚本、文档、代码走查）"
    "效率与覆盖面显著优于人工，但其产出必须经过人工验证才能采信——本次实践中 AI 的两处错误（会话污染、目录误匹配）"
    "均由人工复核发现，且\"预期结果是否符合产品语义\"的判断始终需要人承担最终责任。"
)

AI_CONTRIB = [
    ["成员", "承担任务", "工作量占比"],
    ["甲（王尚清）", "M1/M4 用例与自动化（人工为主）、7 缺陷修复回归、模块一报告 M1/M4 部分", "33.3%"],
    ["乙（待填）", "M2/M5 用例与自动化（待提交后更新）", "33.3%"],
    ["丙", "AI 辅助实践的组织与复核、M3/M6 全部 AI 协作产出（12 用例/8缺陷/三份文档）、模块二报告撰写", "33.3%"],
    ["合计", "", "100%"],
]


def build_report2():
    doc = Document(os.path.join(TPL_DIR, "附录4：测试报告模板（模块二）.docx"))
    fill_cover(doc, "网上书店系统")

    p = replace_body_text(doc, "说明被测软件系统名称", AI_OVERVIEW)
    replace_body_text(doc, "并说明选择方案1", "选择方案2（测试手段含AI），理由见上段。")
    replace_body_text(doc, "方案1：描述被测AI模块", "")
    replace_body_text(doc, "方案2：描述使用的AI工具", "方案2：使用工具为 Claude Code（命令行 AI 编程助手），应用场景覆盖测试环境搭建、测试用例生成、自动化测试脚本开发、缺陷定位与测试文档编写。")

    replace_body_text(doc, "方案1：AI模块的测试用例设计", AI_PROCESS)
    replace_body_text(doc, "方案2：AI的原始输出", "")
    replace_body_text(doc, "方案1：AI相关测试用例清单", AI_RESULTS)

    replace_body_text(doc, "在此说明程序的运行条件", "与模块一测试报告 2.3 节一致（两台开发机：JDK8/Tomcat9.0.121/MySQL5.7.44 与 JDK11/Tomcat9/MySQL8.0；AI 工具运行于本机命令行，无额外服务依赖）。")
    replace_body_text(doc, "并说明测试代码的运行条件", "")

    replace_body_text(doc, "方案1：反思AI系统本身的质量", AI_REFLECT)
    # 列表占位段（方案1的4条）置空
    replace_body_text(doc, "被测AI做对了什么", "")
    replace_body_text(doc, "被测AI做错了什么", "")
    replace_body_text(doc, "测试AI系统的挑战", "")
    replace_body_text(doc, "AI系统能否通过测试证明", "")

    p = replace_body_text(doc, "说明每个成员承担的具体任务", "模块二实践中各成员工作量如下（乙待更新）：")
    insert_table_after(p, doc, AI_CONTRIB)

    doc.save(OUT2)
    print("saved:", OUT2)


if __name__ == "__main__":
    # 仅生成模块一报告（丙个人：M3+M6）；模块二报告见 build_report2（需要时再调用）
    build_report1()
    update_toc(OUT1)
    print("done")
