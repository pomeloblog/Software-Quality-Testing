package com.yj.web;

import com.yj.bean.Book;
import com.yj.bean.Page;
import com.yj.service.BookService;
import com.yj.service.impl.BookServiceImpl;
import com.yj.utils.WebUtils;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;

/**
 * @author yj
 * @create 2020-08-26 15:38
 */
public class ClientBookServlet extends BaseServlet {

    private BookService bookService = new BookServiceImpl();


    protected void page(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        //1、获取请求的参数pageNo和pageSize
        int pageNo = WebUtils.parseInt(req.getParameter("pageNo"),1);
        int pageSize = WebUtils.parseInt(req.getParameter("pageSize"), Page.PAGE_SIZE);

        //2、调用BookService.page(pageNo,pageSize)方法：返回page对象
        Page<Book> page = bookService.page(pageNo,pageSize);
        page.setUrl("client/bookServlet?action=page");
        //3、保存Page对象到request域中
        req.setAttribute("page",page);
        //4、请求转发到page/manager/book_manager.jsp页面
        req.getRequestDispatcher("/pages/client/index.jsp").forward(req,resp);
    }

    protected void pageByNameOrAuthor(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        //1、获取请求的参数pageNo和pageSize、nameorauthor
        int pageNo = WebUtils.parseInt(req.getParameter("pageNo"),1);
        int pageSize = WebUtils.parseInt(req.getParameter("pageSize"), Page.PAGE_SIZE);
       String nameOrAuthor = req.getParameter("nameorauthor");

        //2、调用BookService.page(pageNo,pageSize)方法：返回page对象
        Page<Book> page = bookService.pageByNameOrAuthor(pageNo,pageSize,nameOrAuthor);

        StringBuilder sb = new StringBuilder("client/bookServlet?action=pageByNameOrAuthor");
        if(req.getParameter("nameorauthor")!=null) {
            // 修复 BUG-M3-02：翻页 URL 中的搜索词必须编码，否则空格/&会截断参数、丢失搜索条件
            sb.append("&nameorauthor=").append(URLEncoder.encode(req.getParameter("nameorauthor"), "UTF-8"));
        }
        page.setUrl(sb.toString());
        //3、保存Page对象到request域中
        req.setAttribute("page",page);
        //4、请求转发到page/manager/book_manager.jsp页面
        req.getRequestDispatcher("/pages/client/index.jsp").forward(req,resp);
    }

    protected void pageByPrice(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        //1、获取请求的参数pageNo和pageSize、min、max
        int pageNo = WebUtils.parseInt(req.getParameter("pageNo"),1);
        int pageSize = WebUtils.parseInt(req.getParameter("pageSize"), Page.PAGE_SIZE);
        // 修复 BUG-M3-04：价格字段为 DECIMAL，区间参数改按小数解析（原 parseInt 会把 56.5 静默回退为无筛选）
        double min = WebUtils.parseDouble(req.getParameter("min"),0);
        double max = WebUtils.parseDouble(req.getParameter("max"),99999999);

        // 修复 BUG-M3-03：min>max 时自动交换区间并给出可见提示，而非静默返回空列表
        if (min > max) {
            double t = min; min = max; max = t;
            req.setAttribute("priceMsg","提示：价格区间设置有误（最小值大于最大值），已自动调整区间");
        }

        //2、调用BookService.page(pageNo,pageSize)方法：返回page对象
        Page<Book> page = bookService.pageByPrice(pageNo,pageSize,min,max);

        StringBuilder sb = new StringBuilder("client/bookServlet?action=pageByPrice");
        if(req.getParameter("min")!=null) {
            sb.append("&min=").append(req.getParameter("min"));
        }
        if(req.getParameter("max")!=null) {
            sb.append("&max=").append(req.getParameter("max"));
        }
        page.setUrl(sb.toString());
        //3、保存Page对象到request域中
        req.setAttribute("page",page);
        //4、请求转发到page/manager/book_manager.jsp页面
        req.getRequestDispatcher("/pages/client/index.jsp").forward(req,resp);
    }

    protected void pageOrder(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        //1、获取请求的参数
       //这里直接写死为50
        //2、调用BookService.page(pageNo,pageSize)方法：返回page对象
        Page<Book> page = bookService.pageOrder();
        page.setUrl("client/bookServlet?action=pageOrder");
        //3、保存Page对象到request域中
        req.setAttribute("page",page);
        //4、请求转发到page/manager/book_manager.jsp页面
        req.getRequestDispatcher("/pages/client/top.jsp").forward(req,resp);
    }
}
