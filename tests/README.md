# 自动化测试工程（丙：M3 + M6）

## 范围与用例编号

| 测试类 | 模块 | 用例 |
|---|---|---|
| `M3BookBrowseTest` | M3 图书浏览与检索 | BS-IT-041 ~ BS-IT-046（6 条） |
| `M6ManagerTest` | M6 后台管理 | BS-IT-101 ~ BS-IT-106（6 条） |

方法覆盖：边界值法（041/042/045/104）、等价类法（043/045/102/105）、场景法（101/103/106）、缺陷定向验证（044/046）。

## 环境要求

1. Tomcat 已部署并运行被测应用：<http://localhost:8080/Book/>
2. MySQL 已启动，`book` 库可连接（连接串与被测应用一致，见 `HttpTestUtil.java`：`bookstore/123456`）
3. JDK 8+ 在 PATH 中（javac/java）

依赖 jar 已复制到 `tests/lib/`（junit-4.12、hamcrest-core-1.3、mysql-connector-java-5.1.7），无需 Maven。

## 一键运行

```
tests\run-tests.bat
```

脚本做三件事：编译 `src/*.java` → `classes/`；用 JUnitCore 依次运行两个测试类；打印每个用例结果与汇总。

## 结果如何对应 Excel 清单

- **断言通过 = 用例 OK**；
- **断言失败 = 用例 NG**（失败信息即缺陷证据，Excel `Result` 列填 NG，`Status` 填"不通过"，对应缺陷报告条目）。
- 每个测试方法名带用例号（如 `test042_pageSizeBoundary` ↔ `BS-IT-042`），与 `测试用例/丙-测试用例清单.xlsx` 一一对应。

## 可重复性说明

改库用例均自带数据还原（`try/finally` + `@AfterClass` 兜底清理）：

- `test046`：临时把 id=1 图书销量改为 999，结束还原
- `test102/104`：新增的测试图书按作者标记（`CPTEST`/`CPDEL`）清理
- `test105`：新增用户 `cp_test_01` 用后删除
- `test106`：临时订单 `CPTEST_ORDER_1` 用后删除；`test103` 修改 id=2 图书价格后还原

## 已知环境差异说明

本机运行环境与原始 README 略有不同（对测试无影响）：

| 项 | 原始 README | 本机 |
|---|---|---|
| MySQL | 5.x，root | 8.0，专用账号 `bookstore`（mysql_native_password，兼容项目自带的 5.1.7 驱动） |
| JDK | 8 | 11（编译 `-encoding UTF-8`，测试进程独立于 Tomcat） |
