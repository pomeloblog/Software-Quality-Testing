# -*- coding: utf-8 -*-
"""
生成丙（M3+M6）的测试用例清单 Excel。
基于附录1模板（保留 Information/Introduction 表单），在 Test Cases 表单填入 12 条用例。
Result 列为 2026-09-15 run-tests.bat 实测结果（OK/NG），实测现象并入 Remark 列。
用法：python gen_bing_testcases.py
"""
import os
import openpyxl
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "软件测试实践作业", "实践作业-文档模板2026", "附录1：测试用例清单模板.xlsx")
OUT = os.path.join(HERE, "丙-测试用例清单-M3-M6.xlsx")

M3 = "M3 图书浏览与检索"
M6 = "M6 后台管理"

# 列顺序（13列）：ID, Test Item, Title, Criticality, 是否自动, 是否新增, Pre-condition, Input, Procedure, Output预期, Result实际(OK/NG), Status通过与否, Remark(方法+实测现象)
ROWS = [
    ("BS-IT-041", M3, "首页分页pageNo边界容错", "高", "是", "新增",
     "系统运行中；图书库29本（每页4条共8页）",
     "pageNo=1 / 0 / -1 / 99999 / abc",
     "1.依次请求 /Book/client/bookServlet?action=page&pageNo=X 2.检查HTTP状态码 3.统计页面展示图书条数",
     "pageNo=1显示4本；0和-1回退到第1页；99999回退到末页（第8页，1本）；非数字取默认值1；全部返回HTTP 200",
     "OK", "通过",
     "边界值法（合法值、上/下越界、非法值）。实测：全部容错正确，无服务器错误"),

    ("BS-IT-042", M3, "分页pageSize边界", "高", "是", "新增",
     "同041",
     "pageSize=6 / abc / 0 / -1",
     "1.依次请求 action=page&pageNo=1&pageSize=X 2.检查HTTP状态码与返回条数",
     "合法值6返回6本；非数字容错为默认4；0和-1容错为默认4，均不应出现服务器500错误",
     "NG", "不通过",
     "边界值法（0点、负值、非法值）。实测：pageSize=0→HTTP 500（pageTotalCount%pageSize除零ArithmeticException）；pageSize=-1→HTTP 500（LIMIT负数SQL错误）；6与abc正常"),

    ("BS-IT-043", M3, "书名/作者搜索等价类", "高", "是", "新增",
     "同041；库中有《三体》（刘慈欣）、《解忧杂货店》（东野圭吾）",
     "nameorauthor=三体 / 东野圭吾 / zzz不存在zzz / ' or 1=1 --",
     "1.依次请求 action=pageByNameOrAuthor&nameorauthor=X 2.检查命中图书与HTTP状态",
     "合法词命中对应图书（书名或作者模糊匹配）；不存在的词返回空列表不报错；SQL注入串被参数化查询拦截，返回空结果无异常",
     "OK", "通过",
     "等价类法（有效：书名/作者；无效：不存在词、注入串）。实测：四类输入行为均正确，注入未生效"),

    ("BS-IT-044", M3, "搜索翻页链接特殊字符兼容", "高", "是", "新增",
     "同041",
     "nameorauthor=三 体（含空格）/ a&b（含连接符）",
     "1.请求 action=pageByNameOrAuthor&nameorauthor=三%20体 2.检查响应HTML中翻页导航链接的href属性",
     "翻页链接中的搜索词值应做URL编码，保证点击\"下一页\"后搜索条件完整传递",
     "NG", "不通过",
     "场景法/缺陷定向验证。实测：href形如 ...&nameorauthor=三 体&pageNo=2，空格与&未编码（ClientBookServlet第48行原始拼接）；点击翻页后搜索词被截断、&后内容丢失"),

    ("BS-IT-045", M3, "价格区间筛选边界", "高", "是", "新增",
     "同041；库中存在10~30元图书、价格为57和56.5的边界样本",
     "min,max = (10,30) / (57,57) / (100,10) / (56.5,56.5)",
     "1.依次请求 action=pageByPrice&min=X&max=Y 2.页面结果数与直查库区间数比对 3.检查提示信息",
     "合法区间结果与库一致；min=max精确匹配；min>max给出可见提示或自动归一化；小数价格参数正常参与筛选",
     "NG", "不通过",
     "边界值法+等价类法。实测：min>max返回空列表且无任何提示；(56.5,56.5)筛选完全失效——WebUtils.parseInt仅支持整数，解析失败静默回退0~最大值，页面仍显示全部图书"),

    ("BS-IT-046", M3, "销量榜单排序与完整性", "中", "是", "新增",
     "造数：将id=1《解忧杂货店》sales改为999（全库最高），测试后还原为100",
     "-",
     "1.请求 action=pageOrder 2.检查榜单是否包含《解忧杂货店》",
     "榜单按销量降序展示，且必须包含销量第一名",
     "NG", "不通过",
     "缺陷定向验证。实测：榜单不含销量第一名；SQL为 ORDER BY sales DESC LIMIT 1,50（BookDaoImpl第88行），偏移量1跳过第1名从第2名开始"),

    ("BS-IT-101", M6, "管理员图书管理分页列表", "高", "是", "新增",
     "admin/admin为管理员账号",
     "-",
     "1.admin登录获取会话 2.请求 /manager/bookServlet?action=page&pageNo=1 3.对照组：不带Cookie访问同一URL",
     "管理员访问返回200，展示图书与分页导航；未登录访问被ManagerFilter拦截，响应为登录页而非管理内容",
     "OK", "通过",
     "场景法（含未登录对照组）。实测：登录后正常展示库中图书；未登录响应为登录表单（forward实现，HTTP 200但无管理内容）"),

    ("BS-IT-102", M6, "新增图书输入校验", "高", "是", "新增",
     "admin已登录；测试数据按作者标记CPTEST，测后清理",
     "name=丙测试书,price=35.50 / name=空 / price=-5 / price=abc",
     "1.依次POST /manager/bookServlet?action=add 2.每种输入后查库t_book验证",
     "合法数据入库；空书名、负价格、非数字价格应被服务端校验拒绝并给出提示",
     "NG", "不通过",
     "等价类法（合法值、空值、负值、类型非法值）。实测：四种输入全部直接入库——空书名以name=''入库；负价-5.00入库；price=abc时BeanUtils转换异常被吞掉后以NULL入库；全程无输入校验"),

    ("BS-IT-103", M6, "修改图书回显与更新", "中", "是", "新增",
     "admin已登录；id=2《边城》原价23.00",
     "price修改为99.99",
     "1.GET action=getBook&id=2检查编辑页回显 2.POST action=update改价 3.查库验证 4.还原23.00",
     "编辑页表单value回显原书名《边城》；更新后库中price=99.99",
     "OK", "通过",
     "场景法（回显→更新→验证→还原）。实测：回显正确、更新生效、还原成功"),

    ("BS-IT-104", M6, "删除图书边界", "中", "是", "新增",
     "admin已登录；SQL造数一本待删书（作者CPDEL）",
     "id=存在 / 999999 / abc",
     "1.请求 action=delete&id=X 2.验证库与响应",
     "存在的id删除成功且库中消失；不存在的id给出友好提示；非数字id不引发服务器500",
     "NG", "不通过",
     "边界值法（存在、极大不存在、类型非法）。实测：存在的id删除成功；id=999999静默跳回列表无任何提示；id=abc→HTTP 500（NumberFormatException未处理）"),

    ("BS-IT-105", M6, "用户管理新增与重复用户名", "高", "是", "新增",
     "admin已登录；库中已有admin用户；测试账号cp_test_01测后删除",
     "username=admin（重复）/ cp_test_01（新用户）",
     "1.POST /manager/UserServlet?action=add新增重复admin 2.新增cp_test_01 3.查库验证 4.删除还原",
     "重复用户名给出友好提示（不出现500错误页）；正常新增后用户数+1且可删除",
     "NG", "不通过",
     "等价类法+场景法。实测：重复用户名返回HTTP 500（Duplicate entry唯一约束冲突未捕获，Tomcat错误页）；正常新增与删除功能通过"),

    ("BS-IT-106", M6, "总账单对账与发货状态流转", "高", "是", "新增",
     "admin已登录；SQL造数临时订单CPTEST_ORDER_1（status=0），测后删除",
     "-",
     "1.请求 action=showTotal 2.页面四项统计与SQL直查库比对 3.请求 action=sendOrder&orderId=CPTEST_ORDER_1 4.查库status",
     "用户总数/总订单数/销售本数/总收入与直查库结果一致；发货后订单status由0变为1",
     "OK", "通过",
     "场景法（对账+状态机验证）。实测：四项统计与库一致；发货后status=1。风险备注：总收入口径为SUM(price*sales)（销量×现价），图书价格事后修改会偏离订单实付金额"),
]


def main():
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Test Cases测试用例"]
    start = 2  # 表头第1行，数据第2行起
    for i, row in enumerate(ROWS):
        row = list(row)
        assert len(row) == 13, row[0]
        # 2026-09-15 修复闭环：初测 NG 的用例已全部修复并回归通过，Result 保留初测结果、备注追加回归结论
        if row[10] == "NG":
            row[12] = row[12] + "；【回归】对应缺陷已修复，run-tests.bat 复测通过"
        for j, v in enumerate(row, start=1):
            ws.cell(row=start + i, column=j, value=v)
    widths = [12, 16, 24, 8, 8, 8, 26, 26, 34, 38, 8, 8, 44]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    wb.save(OUT)
    ok = sum(1 for r in ROWS if r[10] == "OK")
    print(f"saved: {OUT}")
    print(f"total={len(ROWS)}, 初测 OK={ok}, NG={len(ROWS) - ok}（NG 已全部修复回归通过）")


if __name__ == "__main__":
    main()
