# -*- coding: utf-8 -*-
"""
将甲/乙/丙三份测试用例清单融合为一份（基于附录1模板）。

输入：
  测试用例清单-M1-M4(甲).xlsx           15条（M1/M4）
  附录1：测试用例清单-乙-M2、M5.xlsx     12条（M2/M5）
  丙-测试用例清单-M3-M6.xlsx             12条（M3/M6）
统一规范：K列=结果代码(OK/POK/NG/NT)，L列=是否通过(通过/不通过)，M列=方法+实测现象。
排序：按模块 M1→M6，模块内按编号数字升序。
用法：python gen_merged_cases.py
"""
import os
import re

import openpyxl
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "软件测试实践作业", "实践作业-文档模板2026", "附录1：测试用例清单模板.xlsx")
OUT = os.path.join(HERE, "附录1：测试用例清单-全组.xlsx")

SOURCES = [
    os.path.join(HERE, "测试用例清单-M1-M4(甲).xlsx"),
    os.path.join(HERE, "附录1：测试用例清单-乙-M2、M5.xlsx"),
    os.path.join(HERE, "丙-测试用例清单-M3-M6.xlsx"),
]

MODULE_ORDER = {"M1": 1, "M2": 2, "M3": 3, "M4": 4, "M5": 5, "M6": 6}


def norm_code(text, status):
    """从 K/L 列内容推断结果代码（OK/POK/NG/NT）。"""
    for cand in (text, status):
        if not cand:
            continue
        s = str(cand).strip()
        if s.startswith(("OK", "POK", "NG", "NT")):
            return s[:3] if s[:3] in ("POK",) else s[:2] if s[:2] in ("OK", "NG") else s[:2]
    # 兜底：L 列是 OK/NG
    if str(status).strip() in ("OK", "POK", "NG", "NT"):
        return str(status).strip()
    return "NT"


def status_zh(code):
    return {"OK": "通过", "POK": "部分通过", "NG": "不通过", "NT": "无法执行"}.get(code, "—")


def extract_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = next(s for s in wb.worksheets if "Test Cases" in s.title)
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        cid = str(r[0] or "").strip()
        if not cid.startswith(("BS", "BS1")):
            continue
        item = str(r[1] or "").strip()
        title = str(r[2] or "").strip()
        crit = str(r[3] or "").strip()
        auto = str(r[4] or "").strip()
        newc = str(r[5] or "").strip()
        pre = r[6]
        inp = r[7]
        proc = r[8]
        output = r[9]
        k_text = str(r[10] or "").strip()   # Result 实际结果
        l_text = str(r[11] or "").strip()   # Status
        remark = str(r[12] or "").strip()

        code = norm_code(k_text, l_text)
        # K 列若为描述性文字（甲乙风格），并入 M 列；已是代码则不动
        k_desc = "" if code in ("OK", "POK", "NG", "NT") and k_text.rstrip("。") == code else ""
        if not k_desc and k_text and not k_text.rstrip("。").startswith(("OK", "POK", "NG", "NT")):
            k_desc = k_text
        elif k_text.startswith(("OK", "POK", "NG", "NT")) and len(k_text) > 4:
            k_desc = re.sub(r"^(OK|POK|NG|NT)[。.，,]?\s*", "", k_text)

        m_parts = [p for p in (remark, ("实测：" + k_desc) if k_desc else "") if p]
        rows.append({
            "cid": cid, "item": item, "title": title, "crit": crit,
            "auto": auto or "是", "newc": newc or "新增", "pre": pre or "",
            "inp": inp or "", "proc": proc or "", "output": output or "",
            "code": code, "status": status_zh(code), "remark": "；".join(m_parts),
            "mod": item[:2] if item[:2] in MODULE_ORDER else "M9",
        })
    return rows


def main():
    all_rows = []
    for p in SOURCES:
        rows = extract_rows(p)
        print(f"{os.path.basename(p)}: {len(rows)}条")
        all_rows.extend(rows)

    all_rows.sort(key=lambda r: (MODULE_ORDER.get(r["mod"], 9),
                                 int(re.search(r"(\d+)\s*$", r["cid"]).group(1)) if re.search(r"(\d+)\s*$", r["cid"]) else 999))

    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Test Cases测试用例"]
    for i, r in enumerate(all_rows):
        vals = [r["cid"], r["item"], r["title"], r["crit"], r["auto"], r["newc"],
                r["pre"], r["inp"], r["proc"], r["output"], r["code"], r["status"], r["remark"]]
        for j, v in enumerate(vals, start=1):
            ws.cell(row=2 + i, column=j, value=v)

    widths = [13, 15, 24, 7, 7, 7, 26, 24, 30, 32, 7, 8, 44]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # Information 表单填基本信息
    if "Information文档信息" in wb.sheetnames:
        info = wb["Information文档信息"]
        filled = 0
        for row in info.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "项目名称" in cell.value and filled == 0:
                    info.cell(row=cell.row, column=cell.column + 1, value="网上书店系统")
                    filled = 1

    wb.save(OUT)

    # 汇总统计
    from collections import Counter
    mods = Counter(r["mod"] for r in all_rows)
    codes = Counter(r["code"] for r in all_rows)
    autos = sum(1 for r in all_rows if "是" in r["auto"])
    print(f"\nmerged -> {OUT}")
    print("模块分布:", dict(sorted(mods.items())))
    print("结果分布:", dict(codes), f"| 总计 {len(all_rows)} 条，自动化 {autos} 条")


if __name__ == "__main__":
    main()
