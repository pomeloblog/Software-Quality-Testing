package com.yj.test;

import com.yj.test.common.HttpTestClient;
import com.yj.test.common.TestDb;
import com.yj.test.m1.M1UserAuthTest;
import com.yj.test.m4.M4CartTest;
import org.junit.runner.Description;
import org.junit.runner.JUnitCore;
import org.junit.runner.Result;
import org.junit.runner.notification.Failure;
import org.junit.runner.notification.RunListener;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 自动化测试统一入口（一键运行）。
 *
 * <p>用 {@link JUnitCore} 执行全部用例，并按**用例编号**输出逐条结果与汇总统计，
 * 输出可直接作为测试报告中的"执行日志"。
 *
 * <p>注意：标有 {@code _defect_} 的用例，断言写的是正确行为，应用存在缺陷时**本就应该失败**
 * （对应用例清单中的 NG）。因此"运行成功"的标准是**全部用例都被执行并给出判定**，
 * 而不是"没有失败"——这是缺陷驱动测试的正常现象。
 *
 * @author 王尚清
 */
public class TestRunner {

    private static final String BASE = "http://localhost:8080/Book";

    /** 模块名 -> 测试类。新增模块时在此登记即可。 */
    private static final Map<String, Class<?>> SUITES = new LinkedHashMap<String, Class<?>>();

    static {
        SUITES.put("M1 用户认证", M1UserAuthTest.class);
        SUITES.put("M4 购物车", M4CartTest.class);
    }

    /** bsIt001_xxx / bsUt061_xxx */
    private static final Pattern CASE_ID = Pattern.compile("^(bsIt|bsUt)(\\d{3})_(.*)$");

    private static final Map<String, String> STATUS = new TreeMap<String, String>();
    private static final Map<String, String> TITLE = new TreeMap<String, String>();
    private static final Map<String, String> REASON = new TreeMap<String, String>();
    private static final Map<String, String> SUITE_OF = new TreeMap<String, String>();

    public static void main(String[] args) {
        banner();
        preconditions();

        System.out.println("---------------- 逐条结果 ----------------");
        JUnitCore core = new JUnitCore();
        core.addListener(new Collector());
        for (Map.Entry<String, Class<?>> entry : SUITES.entrySet()) {
            Result result = core.run(entry.getValue());
            if (result.getRunCount() == 0) {
                System.out.println("  !! " + entry.getKey() + " 未执行到任何用例，请检查测试类 "
                        + entry.getValue().getName());
            }
        }

        int pass = 0;
        for (Map.Entry<String, String> e : STATUS.entrySet()) {
            String id = e.getKey();
            String status = e.getValue();
            if ("PASS".equals(status)) {
                pass++;
            }
            System.out.printf("[%s] %-11s %s%n", status, id, TITLE.get(id));
            String reason = REASON.get(id);
            if ("FAIL".equals(status) && reason != null) {
                System.out.println("          └─ " + reason);
            }
        }

        summary(pass);
    }

    /** 收集每条用例的结果，并按方法名还原用例编号。 */
    private static class Collector extends RunListener {
        @Override
        public void testFinished(Description d) {
            String id = caseId(d.getMethodName());
            if (id == null) {
                return;
            }
            // JUnit 的回调顺序是 testStarted -> testFailure -> testFinished，
            // 因此这里必须先判断是否已判定为 FAIL，否则会把失败覆盖成通过。
            if ("FAIL".equals(STATUS.get(id))) {
                return;
            }
            record(d, "PASS", null);
        }

        @Override
        public void testFailure(Failure f) {
            record(f.getDescription(), "FAIL", firstLine(f.getMessage()));
        }

        @Override
        public void testAssumptionFailure(Failure f) {
            record(f.getDescription(), "SKIP", firstLine(f.getMessage()));
        }

        @Override
        public void testIgnored(Description d) {
            record(d, "SKIP", "用例被 @Ignore 忽略");
        }

        private void record(Description d, String status, String reason) {
            String id = caseId(d.getMethodName());
            if (id == null) {
                return;   // 不符合命名约定的方法不纳入统计
            }
            STATUS.put(id, status);
            TITLE.put(id, title(d.getMethodName()));
            SUITE_OF.put(id, d.getTestClass().getSimpleName());
            if (reason != null) {
                REASON.put(id, reason);
            }
        }
    }

    private static void summary(int pass) {
        int total = STATUS.size();
        int fail = 0;
        int skip = 0;
        for (String s : STATUS.values()) {
            if ("FAIL".equals(s)) {
                fail++;
            } else if ("SKIP".equals(s)) {
                skip++;
            }
        }

        System.out.println();
        System.out.println("---------------- 汇总 ----------------");
        for (String suiteName : SUITES.keySet()) {
            String simple = SUITES.get(suiteName).getSimpleName();
            int n = 0;
            for (String s : SUITE_OF.values()) {
                if (simple.equals(s)) {
                    n++;
                }
            }
            System.out.printf("  %-14s %d 条用例%n", suiteName, n);
        }
        System.out.printf("  用例总数: %d    通过: %d    失败: %d    跳过: %d%n", total, pass, fail, skip);
        System.out.printf("  通过率: %.1f%%%n", total == 0 ? 0.0 : (pass * 100.0 / total));
        System.out.println();
        if (fail > 0) {
            System.out.println("  说明: 未通过的用例，其断言写的是「正确行为」，用于暴露尚未修复的缺陷，");
            System.out.println("        失败即代表缺陷仍然存在（对应用例清单中的 NG）。");
        } else {
            System.out.println("  说明: 全部用例通过。原缺陷用例已保留为回归测试，");
            System.out.println("        若相关缺陷回归，会立即重新失败。");
        }
        System.out.println("========================================");
    }

    private static void banner() {
        System.out.println("========================================");
        System.out.println("  网上书店系统 · 自动化测试");
        System.out.println("  被测模块: M1 用户认证 / M4 购物车");
        System.out.println("  作者: 王尚清");
        System.out.println("========================================");
    }

    private static void preconditions() {
        HttpTestClient probe = new HttpTestClient(BASE);
        boolean appUp = probe.serverReachable();
        boolean dbUp = TestDb.available();

        System.out.println("---------------- 前置检查 ----------------");
        System.out.println("  MySQL:  " + (dbUp ? "可连接" : "**不可连接**"));
        System.out.println("  应用:   " + BASE + "  " + (appUp ? "可访问" : "**不可访问**"));
        if (!dbUp || !appUp) {
            System.out.println();
            System.out.println("  !! 前置条件不满足，集成测试将大面积失败，结果不具备参考价值。");
            System.out.println("  !! 请先启动 MySQL 与 Tomcat，再重新运行本测试。");
        }
        System.out.println();
    }

    /** bsIt001_xxx -> BS-IT-001 */
    static String caseId(String methodName) {
        Matcher m = CASE_ID.matcher(methodName);
        if (!m.matches()) {
            return null;
        }
        String kind = "bsIt".equals(m.group(1)) ? "IT" : "UT";
        return "BS-" + kind + "-" + m.group(2);
    }

    /** bsIt001_loginSucceeds -> loginSucceeds */
    static String title(String methodName) {
        Matcher m = CASE_ID.matcher(methodName);
        return m.matches() ? m.group(3) : methodName;
    }

    private static String firstLine(String message) {
        if (message == null) {
            return null;
        }
        int nl = message.indexOf('\n');
        return nl < 0 ? message : message.substring(0, nl);
    }
}
