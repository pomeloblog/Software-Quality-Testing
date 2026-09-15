import org.junit.runner.Description;
import org.junit.runner.JUnitCore;
import org.junit.runner.Result;
import org.junit.runner.notification.Failure;
import org.junit.runner.notification.RunListener;

/**
 * 一键测试运行入口：替代裸 JUnitCore，逐条打印每个用例的执行情况。
 * 输出格式与 Excel 用例清单（Result 列）一一对应：
 *   BS-IT-041  test041_pageNoBoundary ... 通过 (12 ms)
 * bat 中由 run-tests.bat 调用，退出码 0=全部通过 / 1=存在失败。
 */
public class RunAllTests {

    /** 逐用例进度监听：started 记时，failure 标记，finished 打印一行结果。 */
    public static class ProgressListener extends RunListener {
        private long t0;
        private boolean failed;

        @Override
        public void testStarted(Description d) {
            failed = false;
            t0 = System.currentTimeMillis();
        }

        @Override
        public void testFailure(Failure f) {
            failed = true;
        }

        @Override
        public void testFinished(Description d) {
            long ms = System.currentTimeMillis() - t0;
            // test041_pageNoBoundary -> BS-IT-041（与用例清单编号对应）
            String id = d.getMethodName().replaceFirst("^test(\\d+)_.*", "BS-IT-$1");
            String verdict = failed ? "不通过 NG" : "通过 OK";
            System.out.println(String.format("  %-10s %-36s %-9s (%d ms)",
                    id, d.getMethodName(), verdict, ms));
        }
    }

    public static void main(String[] args) {
        JUnitCore core = new JUnitCore();
        core.addListener(new ProgressListener());

        System.out.println("== M3 图书浏览与检索 (BS-IT-041 ~ 046) ==");
        Result r3 = core.run(M3BookBrowseTest.class);
        sub(r3);

        System.out.println("== M6 后台管理 (BS-IT-101 ~ 106) ==");
        Result r6 = core.run(M6ManagerTest.class);
        sub(r6);

        int total = r3.getRunCount() + r6.getRunCount();
        int fail = r3.getFailureCount() + r6.getFailureCount();
        System.out.println("==================== 汇总 ====================");
        System.out.println("  总用例: " + total + "    通过: " + (total - fail)
                + "    不通过: " + fail
                + "    通过率: " + (total == 0 ? 0 : (total - fail) * 100 / total) + "%");
        if (fail > 0) {
            System.out.println("  ---- 失败明细（对应 Excel 清单 NG 行）----");
            printFailures("M3", r3);
            printFailures("M6", r6);
        }
        System.exit(fail == 0 ? 0 : 1);
    }

    private static void sub(Result r) {
        System.out.println(String.format("  -- 小计: %d/%d 通过 --",
                r.getRunCount() - r.getFailureCount(), r.getRunCount()));
    }

    private static void printFailures(String tag, Result r) {
        for (Failure f : r.getFailures()) {
            String msg = f.getMessage() == null ? "(无断言信息)" : f.getMessage().replace("\r", " ").replace("\n", " ; ");
            System.out.println(String.format("  [%s] %s : %s", tag, f.getDescription().getMethodName(), msg));
        }
    }
}
