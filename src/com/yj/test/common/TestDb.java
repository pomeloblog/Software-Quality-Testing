package com.yj.test.common;

import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Properties;

/**
 * 测试用数据库直连工具。
 *
 * <p>刻意**不复用**被测代码的 {@link com.yj.utils.JDBCUtils}：后者把连接放在 ThreadLocal 里
 * 并由 TransactionFilter 负责提交，测试中直接使用会与其事务模型耦合。
 * 这里用 DriverManager 独立开连接，保证断言看到的是数据库的真实落库结果。
 *
 * <p>连接参数读自 classpath 下的 jdbc.properties（与生产配置同一份）。
 *
 * @author 王尚清
 */
public class TestDb {

    private static final String URL;
    private static final String USER;
    private static final String PASSWORD;

    static {
        try {
            Properties p = new Properties();
            InputStream in = TestDb.class.getClassLoader().getResourceAsStream("jdbc.properties");
            if (in == null) {
                throw new IllegalStateException(
                        "classpath 下找不到 jdbc.properties —— 请确认构建时已把 src/jdbc.properties 复制到 classes 目录");
            }
            p.load(in);
            in.close();
            URL = p.getProperty("url");
            USER = p.getProperty("username");
            PASSWORD = p.getProperty("password");
            Class.forName(p.getProperty("driverClassName", "com.mysql.jdbc.Driver"));
        } catch (Exception e) {
            throw new ExceptionInInitializerError(e);
        }
    }

    public static Connection open() throws SQLException {
        return DriverManager.getConnection(URL, USER, PASSWORD);
    }

    /** 数据库中是否存在该用户名。 */
    public static int countUser(String username) {
        Connection conn = null;
        PreparedStatement ps = null;
        ResultSet rs = null;
        try {
            conn = open();
            ps = conn.prepareStatement("select count(*) from t_user where username = ?");
            ps.setString(1, username);
            rs = ps.executeQuery();
            return rs.next() ? rs.getInt(1) : 0;
        } catch (SQLException e) {
            throw new RuntimeException("查询用户失败: " + username, e);
        } finally {
            close(rs, ps, conn);
        }
    }

    /** 删除测试用户，用于用例前置准备与善后清理。 */
    public static void deleteUser(String username) {
        Connection conn = null;
        PreparedStatement ps = null;
        try {
            conn = open();
            ps = conn.prepareStatement("delete from t_user where username = ?");
            ps.setString(1, username);
            ps.executeUpdate();
        } catch (SQLException e) {
            throw new RuntimeException("删除用户失败: " + username, e);
        } finally {
            close(null, ps, conn);
        }
    }

    /** 某本图书的库存。 */
    public static int bookStock(int bookId) {
        Connection conn = null;
        PreparedStatement ps = null;
        ResultSet rs = null;
        try {
            conn = open();
            ps = conn.prepareStatement("select stock from t_book where id = ?");
            ps.setInt(1, bookId);
            rs = ps.executeQuery();
            return rs.next() ? rs.getInt(1) : -1;
        } catch (SQLException e) {
            throw new RuntimeException("查询库存失败: id=" + bookId, e);
        } finally {
            close(rs, ps, conn);
        }
    }

    /** 数据库是否可连通，用于运行时前置检查。 */
    public static boolean available() {
        try {
            Connection conn = open();
            conn.close();
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private static void close(ResultSet rs, PreparedStatement ps, Connection conn) {
        try {
            if (rs != null) {
                rs.close();
            }
        } catch (SQLException ignored) {
        }
        try {
            if (ps != null) {
                ps.close();
            }
        } catch (SQLException ignored) {
        }
        try {
            if (conn != null) {
                conn.close();
            }
        } catch (SQLException ignored) {
        }
    }
}
