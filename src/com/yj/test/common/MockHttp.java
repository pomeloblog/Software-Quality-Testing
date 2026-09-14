package com.yj.test.common;

import javax.servlet.RequestDispatcher;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 用 JDK 动态代理构造 {@link HttpServletRequest} / {@link HttpSession} / {@link RequestDispatcher}
 * 的桩（Stub）对象，使 Servlet 可以在**不启动容器**的情况下被直接调用。
 *
 * <p>为什么需要它：注册流程被验证码保护（必须有正确的会话验证码才能走到"密码一致性校验"分支），
 * 走 HTTP 无法构造该前置条件。用桩把验证码直接注入会话，即可到达目标分支，
 * 从而对 BS-UT-006 做**动态**验证，而不是仅靠阅读代码。
 *
 * <p>属于白盒单元测试手段，对应测试报告"如何处理外部依赖（Mock/Stub）"一节。
 *
 * @author 王尚清
 */
public class MockHttp {

    /** kaptcha 把验证码存进会话时使用的属性名。 */
    public static final String KAPTCHA_SESSION_KEY = "KAPTCHA_SESSION_KEY";

    private final Map<String, String> params = new LinkedHashMap<String, String>();
    /** 会话属性，测试可直接读写（例如注入验证码）。 */
    public final Map<String, Object> sessionAttrs = new LinkedHashMap<String, Object>();
    /** 请求属性，测试可读取 Servlet 放进去的提示信息。 */
    public final Map<String, Object> requestAttrs = new LinkedHashMap<String, Object>();
    /** 最后一次 request.getRequestDispatcher(path).forward(...) 的目标页面。 */
    public String forwardedTo;
    /** 最后一次 response.sendRedirect(location) 的目标地址。 */
    public String redirectedTo;
    private final StringWriter responseBuffer = new StringWriter();

    public MockHttp() {
    }

    public MockHttp(Map<String, String> params) {
        if (params != null) {
            this.params.putAll(params);
        }
    }

    public MockHttp put(String key, String value) {
        params.put(key, value);
        return this;
    }

    /** 写入验证码，模拟"用户已从验证码图片取得验证码"。 */
    public MockHttp withCaptcha(String code) {
        sessionAttrs.put(KAPTCHA_SESSION_KEY, code);
        return this;
    }

    /** response.getWriter() 累计写出的内容。 */
    public String responseBody() {
        return responseBuffer.toString();
    }

    public HttpServletRequest request() {
        return (HttpServletRequest) proxy(HttpServletRequest.class, new InvocationHandler() {
            @Override
            public Object invoke(Object p, Method m, Object[] args) throws Throwable {
                String n = m.getName();
                if ("getSession".equals(n)) {
                    return session();
                }
                if ("getParameter".equals(n)) {
                    return params.get(args[0]);
                }
                if ("getParameterMap".equals(n)) {
                    Map<String, String[]> map = new LinkedHashMap<String, String[]>();
                    for (Map.Entry<String, String> e : params.entrySet()) {
                        map.put(e.getKey(), new String[]{e.getValue()});
                    }
                    return map;
                }
                if ("setCharacterEncoding".equals(n)) {
                    return null;
                }
                if ("setAttribute".equals(n)) {
                    requestAttrs.put((String) args[0], args[1]);
                    return null;
                }
                if ("getAttribute".equals(n)) {
                    return requestAttrs.get(args[0]);
                }
                if ("removeAttribute".equals(n)) {
                    requestAttrs.remove(args[0]);
                    return null;
                }
                if ("getRequestDispatcher".equals(n)) {
                    return dispatcher((String) args[0]);
                }
                if ("getContextPath".equals(n)) {
                    return "";
                }
                if ("getMethod".equals(n)) {
                    return "POST";
                }
                if ("getProtocol".equals(n)) {
                    return "HTTP/1.1";
                }
                return fallback(p, m, args);
            }
        });
    }

    public HttpServletResponse response() {
        return (HttpServletResponse) proxy(HttpServletResponse.class, new InvocationHandler() {
            @Override
            public Object invoke(Object p, Method m, Object[] args) throws Throwable {
                String n = m.getName();
                if ("getWriter".equals(n)) {
                    return new PrintWriter(responseBuffer);
                }
                if ("setContentType".equals(n) || "setCharacterEncoding".equals(n)) {
                    return null;
                }
                if ("sendRedirect".equals(n)) {
                    redirectedTo = (String) args[0];
                    return null;
                }
                return fallback(p, m, args);
            }
        });
    }

    public HttpSession session() {
        return (HttpSession) proxy(HttpSession.class, new InvocationHandler() {
            @Override
            public Object invoke(Object p, Method m, Object[] args) throws Throwable {
                String n = m.getName();
                if ("getAttribute".equals(n)) {
                    return sessionAttrs.get(args[0]);
                }
                if ("setAttribute".equals(n)) {
                    sessionAttrs.put((String) args[0], args[1]);
                    return null;
                }
                if ("removeAttribute".equals(n)) {
                    sessionAttrs.remove(args[0]);
                    return null;
                }
                if ("invalidate".equals(n)) {
                    sessionAttrs.clear();
                    return null;
                }
                if ("getId".equals(n)) {
                    return "MOCK-SESSION";
                }
                return fallback(p, m, args);
            }
        });
    }

    private RequestDispatcher dispatcher(final String path) {
        return (RequestDispatcher) proxy(RequestDispatcher.class, new InvocationHandler() {
            @Override
            public Object invoke(Object p, Method m, Object[] args) throws Throwable {
                String n = m.getName();
                if ("forward".equals(n) || "include".equals(n)) {
                    if ("forward".equals(n)) {
                        forwardedTo = path;
                    }
                    return null;
                }
                return fallback(p, m, args);
            }
        });
    }

    private Object proxy(Class<?> type, InvocationHandler handler) {
        return Proxy.newProxyInstance(type.getClassLoader(), new Class<?>[]{type}, handler);
    }

    /** 未被显式处理的方法：Object 的三个基础方法给出合理实现，其余按返回值类型给默认值。 */
    private static Object fallback(Object proxy, Method m, Object[] args) {
        String n = m.getName();
        if ("toString".equals(n)) {
            return "MockHttpProxy";
        }
        if ("hashCode".equals(n)) {
            return System.identityHashCode(proxy);
        }
        if ("equals".equals(n)) {
            return proxy == (args == null || args.length == 0 ? null : args[0]);
        }
        Class<?> rt = m.getReturnType();
        if (!rt.isPrimitive() || void.class.equals(rt)) {
            return null;
        }
        if (boolean.class.equals(rt)) {
            return Boolean.FALSE;
        }
        if (char.class.equals(rt)) {
            return (char) 0;
        }
        if (byte.class.equals(rt)) {
            return (byte) 0;
        }
        if (short.class.equals(rt)) {
            return (short) 0;
        }
        if (int.class.equals(rt)) {
            return 0;
        }
        if (long.class.equals(rt)) {
            return 0L;
        }
        if (float.class.equals(rt)) {
            return 0f;
        }
        if (double.class.equals(rt)) {
            return 0d;
        }
        return null;
    }

    /** 便于断言时打印 Servlet 投向的页面。 */
    public String describe() {
        StringBuilder sb = new StringBuilder();
        if (forwardedTo != null) {
            sb.append("forward -> ").append(forwardedTo);
        }
        if (redirectedTo != null) {
            if (sb.length() > 0) {
                sb.append(", ");
            }
            sb.append("redirect -> ").append(redirectedTo);
        }
        if (requestAttrs.containsKey("msg")) {
            if (sb.length() > 0) {
                sb.append(", ");
            }
            sb.append("msg=").append(requestAttrs.get("msg"));
        }
        return sb.toString();
    }
}
