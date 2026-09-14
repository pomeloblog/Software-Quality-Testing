# 网上书店系统 · 自动化测试工程说明

> 课程：软件测试与质量保证实践　阶段：模块一（测试基础实践）
> 被测模块：**M1 用户认证**、**M4 购物车**　作者：王尚清
> 被测对象：网上书店系统（JavaWeb MVC）

---

## 1. 测试范围

| 项目 | 说明 |
|---|---|
| **本次测试覆盖** | **M1 用户认证**：登录（正确/密码错/用户不存在）、验证码校验、注册、用户名查重<br>**M4 购物车**：加入购物车、重复加购累加、修改数量、删除、清空、总金额计算 |
| **本次测试不覆盖** | ① M1 的**注销**与**个人信息修改**——优先级较低，用例额度分配给了缺陷更集中的路径<br>② M4 的**页面样式与交互视觉效果**——非本阶段测试范畴<br>③ 图书浏览检索(M3)、订单(M5)、后台管理(M6)——由小组其他成员负责<br>④ 性能、并发、兼容性测试——本阶段不涉及 |
| **不覆盖原因** | 环境限制 / 优先级排序 / 非本模块职责 / 阶段范围外 |

**测试方法**：等价类划分、边界值分析、场景法（共 3 种，满足"≥2 种"要求）；
白盒手段上，单元测试直接调用被测类，其余用例通过 HTTP 打真实 Servlet；
其中 **4 条用例（BS-IT-001 ~ 003、BS-UT-006）使用动态代理桩（Stub）** 突破验证码门槛。

---

## 2. 目录结构

```
├── run-tests.bat                          一键运行脚本（编译 + 执行 + 输出报告）
├── build/                                 编译输出（构建产物，不入版本库）
└── src/com/yj/test/
    ├── common/                            测试基础设施
    │   ├── HttpTestClient.java            集成测试 HTTP 客户端（自动管理 JSESSIONID 会话）
    │   ├── MockHttp.java                  Servlet 单元测试桩（JDK 动态代理实现）
    │   └── TestDb.java                    测试直连数据库，用于断言真实落库结果
    ├── m1/
    │   └── M1UserAuthTest.java            M1 用户认证 7 条用例
    ├── m4/
    │   └── M4CartTest.java                M4 购物车 8 条用例
    └── TestRunner.java                    一键运行入口，按用例编号输出结果与汇总
```

**设计要点**

- `HttpTestClient` 自带 Cookie 管理——购物车存在 session 中，必须靠它维持会话状态。
- `TestDb` **刻意不复用**被测代码的 `JDBCUtils`：后者把连接放在 ThreadLocal 里并由过滤器提交，
  测试中直接使用会与其事务模型耦合。这里独立开连接，保证断言看到的是数据库的真实结果。
- `MockHttp` 用 `java.lang.reflect.Proxy` 构造 `HttpServletRequest` / `HttpSession` / `RequestDispatcher`
  的桩对象，使 Servlet 能在**不启动容器**的情况下被调用。它存在的根本原因是**验证码**：
  注册与登录都受验证码保护，而自动化脚本无法识别验证码图片（这正是验证码的设计目的），
  因此 BS-IT-001 ~ 003、BS-UT-006 改用桩向会话注入验证码，**仅替换容器的传输层**，
  断言的仍是「Servlet → Service → DAO → MySQL」全链路。
- **桩驱动的用例用真实的 `TransactionFilter` 包住被调用的 Servlet**。这一点很关键：被测系统把事务边界
  交给了过滤器（它负责 commit / rollback），`BaseServlet` 内部的 DAO 拿到的是 `autoCommit=false` 的连接。
  若绕过过滤器直接调用，写操作既不会提交（断言永远看不到数据），又会因未提交的 INSERT 持有行锁
  拖住后续清理语句。因此测试用真实过滤器驱动，事务提交/回滚完全由被测代码自己完成。

---

## 3. 环境要求

| 项 | 版本 / 说明 |
|---|---|
| 操作系统 | Windows 11 |
| JDK | **1.8.0_504**（Temurin）——必须用 1.8，见下方警告 |
| 应用服务器 | Tomcat **9.0.121**，context path 固定为 `/Book` |
| 数据库 | MySQL **5.7.44**，库名 `book` |
| 测试框架 | JUnit **4.12** + hamcrest-core **1.3**（已在 `web/WEB-INF/lib/`） |
| 数据库账号 | `root` / `12345`（见 `src/jdbc.properties`） |
| 系统账号 | 管理员 `admin` / `admin` |

> ⚠️ **不要用现代 JDK 编译本项目**。项目 `.idea/misc.xml` 里写着 `languageLevel="JDK_22"`，
> 但 Tomcat 运行在 Java 8 上。用高版本 JDK 编译出的 class 是 version 66.0，Java 8 只认 52.0，
> 会导致 `UnsupportedClassVersionError`，应用启动时报
> **"一个或多个筛选器启动失败 / Context[/Book] 启动失败"**。

---

## 4. 一键运行

### 前置：启动 MySQL 与 Tomcat

`run-tests.bat` **只跑测试，不负责启动应用**。请先确保应用已在运行：

```bat
:: 窗口 1 —— MySQL
E:\网上书店项目依赖\runtime\mysql-5.7.44-winx64\bin\mysqld.exe --datadir="E:\网上书店项目依赖\runtime\mysql-data" --port=3306 --console

:: 窗口 2 —— Tomcat
set JAVA_HOME=E:\网上书店项目依赖\runtime\jdk8u504-b01
set CATALINA_HOME=E:\网上书店项目依赖\runtime\apache-tomcat-9.0.121
E:\网上书店项目依赖\runtime\apache-tomcat-9.0.121\bin\catalina.bat run
```

确认 <http://localhost:8080/Book/> 能打开首页。

### 执行测试

双击项目根目录的 **`run-tests.bat`**（或在项目根目录执行 `run-tests.bat`）。

脚本依次完成：检查运行环境 → 准备依赖 jar → 生成源文件清单 → 编译被测代码与测试代码 → 执行全部用例并输出结果。

**若 JDK / Tomcat 不在上述路径**，修改 `run-tests.bat` 顶部这两行即可：

```bat
set "JAVA_HOME=E:\网上书店项目依赖\runtime\jdk8u504-b01"
set "TOMCAT_HOME=E:\网上书店项目依赖\runtime\apache-tomcat-9.0.121"
```

### 输出说明

脚本会先做前置检查（MySQL 是否可连、应用是否可访问），再逐条输出：

```
[PASS] BS-IT-001   loginSucceedsWithCorrectPassword
[PASS] BS-IT-004   loginRejectedWithoutCaptcha
[PASS] BS-IT-065   negativeCountIsRejected
...
[FAIL] BS-XX-0NN   someCaseName
          └─ 断言的说明信息（含实际观测到的现象）

---------------- 汇总 ----------------
  用例总数: 15    通过: 15    失败: 0    跳过: 0
  通过率: 100.0%
```

该输出可直接作为测试报告中的**执行日志**。出现 `[FAIL]` 说明仍有未修复的缺陷。

---

## 5. 用例与代码对应关系

### M1 用户认证（7 条）

| 用例编号 | 测试方法 | 类型 | 说明 |
|---|---|---|---|
| BS-IT-001 | `bsIt001_loginSucceedsWithCorrectPassword` | 桩驱动 | 等价类-有效 |
| BS-IT-002 | `bsIt002_loginFailsWithWrongPassword` | 桩驱动 | 等价类-无效 |
| BS-IT-003 | `bsIt003_loginFailsWithUnknownUsername` | 桩驱动 | 等价类-无效 |
| BS-IT-004 | `bsIt004_loginRejectedWithoutCaptcha` | 集成 | 不携带验证码应被拒绝 |
| BS-IT-005 | `bsIt005_registRejectedWithoutCaptcha` | 集成 | 场景法 |
| BS-UT-006 | `bsUt006_mismatchedPasswordIsRejected` | 桩驱动 | 两次密码不一致应被拒绝 |
| BS-IT-007 | `bsIt007_ajaxExistsUsername` | 集成 | 等价类 |

> BS-IT-001 ~ 003 之所以是"桩驱动"而非走 HTTP：登录已补齐验证码校验，
> 而自动化脚本无法识别验证码图片，故改用桩注入会话验证码（详见 §2 设计要点）。

### M4 购物车（8 条）

| 用例编号 | 测试方法 | 类型 | 说明 |
|---|---|---|---|
| BS-UT-061 | `bsUt061_repeatAddAccumulatesCountAndPrice` | 单元 | 等价类 |
| BS-UT-062 | `bsUt062_totalPriceUsesBigDecimalPrecision` | 单元 | 数据精度 |
| BS-IT-063 | `bsIt063_addNonexistentBookIsHandled` | 集成 | 加购不存在图书 ID 应被友好处理 |
| BS-IT-064 | `bsIt064_addNonNumericIdIsHandled` | 集成 | 加购非数字 ID 应被友好处理 |
| BS-IT-065 | `bsIt065_negativeCountIsRejected` | 集成 | 负数数量应被拒绝并保留原值 |
| BS-IT-066 | `bsIt066_nonNumericCountIsRejected` | 集成 | 非法数量应被拒绝并保留原值 |
| BS-IT-067 | `bsIt067_emptyCartDeleteStillResponds` | 集成 | 空购物车删除仍应给出响应 |
| BS-IT-068 | `bsIt068_clearItemEmptiesCart` | 集成 | 场景法 |

> 063 ~ 067 原为暴露缺陷的用例（初测判定 NG），缺陷修复后转为**回归测试**：
> 断言写的是修复后的正确行为，若缺陷回归会立即重新失败。

---

## 6. 缺陷与修复

**初测结果：15 条用例，通过 8 条、未通过 7 条，通过率 53.3%。**
未通过的 7 条正好对应发现并确认的 **7 个有效缺陷**。

面向缺陷的用例，其**断言写的是正确行为**（例如"登录请求未携带验证码时应当被拒绝"）。
应用存在该缺陷时断言自然不成立，于是判定为不通过 —— 正好对应用例清单中 `Result` 列为 **NG** 的条目。
因此本工程"运行成功"的标准是**全部用例都被执行并给出判定**，而不是"没有失败"。

**7 个缺陷全部已修复**（改动与验证详见《测试缺陷报告》3.5.5 节）：

| 缺陷编号 | 缺陷 | 位置 | 级别 | 状态 |
|---|---|---|---|---|
| DEF-001 | 登录不校验验证码，可绕过做暴力破解 | `UserServlet.login()` | 高（安全） | 已修复 |
| DEF-002 | 商品数量可改为负数，产生负数量、负金额 | `Cart.updateCount()` 无下界校验 | 高（数据） | 已修复 |
| DEF-003 | 非法数量被静默改成 1，用户输入被篡改 | `CartServlet.updateCount()` | 高（数据） | 已修复 |
| DEF-004 | 注册不校验两次密码一致性 | `UserServlet.regist()` | 中 | 已修复 |
| DEF-005 | 加购不存在的图书 ID 触发空指针返回 500 | `CartServlet.ajaxAddItem()` | 中 | 已修复 |
| DEF-006 | 加购非数字 ID 同样返回 500 | `WebUtils.parseInt()` 静默取默认值 | 中 | 已修复 |
| DEF-007 | 空购物车时删除/改数量返回空响应，页面卡死 | `CartServlet.deleteItem()` / `updateCount()` | 中 | 已修复 |

> DEF-005 与 DEF-006 触发位置相同（同一行空指针），但成因不同：前者缺「图书存在性校验」，
> 后者缺「参数格式校验」，修复动作也不同，故分别记录。

**回归结果：修复后重跑全部 15 条用例，15 条全部通过，通过率 100%**，无新增缺陷或回归。
标有"原缺陷"语义的用例已作为**回归测试**保留 —— 若缺陷回归，会立即重新失败。

| 轮次 | 用例总数 | 通过 | 未通过 | 通过率 |
|---|---|---|---|---|
| 初测（缺陷发现） | 15 | 8 | 7 | 53.3% |
| 回归（修复后） | 15 | 15 | 0 | 100.0% |

---

## 7. 如何新增用例

1. 在对应模块的测试类中新增方法，**方法名必须符合** `bsIt001_描述` 或 `bsUt061_描述` 的格式
   （小写 `bs` + `It`/`Ut` + 三位用例编号 + 下划线 + 描述），`TestRunner` 会自动把它还原成
   用例编号（`BS-IT-001`）并纳入统计。
2. 若新增测试类（例如其他成员负责的 M3/M5/M6），在 `TestRunner.SUITES` 中登记模块名与类名即可。
3. **测试须可重复运行**：改动了数据库的用例，请在 `@Before` 准备数据、`@After` 清理还原
   （例如 `TestDb.deleteUser(...)`）。当前全部用例已做到运行前后数据库状态不变。

---

## 8. 其他说明

- **自动化测试本身不修改被测代码**；缺陷修复是独立于测试的一步，本工程中缺陷确已修复，改动清单与验证见《测试缺陷报告》3.5.5 节。
- 用例范围内的所有断言都基于**实际执行结果**，不包含仅凭阅读代码得出的结论
  （BS-UT-006 最初由代码审查发现线索，最终已通过桩注入**动态复现**）。
- `build/` 为编译产物，不应提交到版本库。
