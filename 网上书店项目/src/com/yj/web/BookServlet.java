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
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.util.List;

/**
 * @author yj
 * @create 2020-08-24 15:25
 */
public class BookServlet extends BaseServlet {

    private BookService bookService = new BookServiceImpl();

    /**
     * 图书输入校验。修复 BUG-M6-01：此前空书名/负价格/非法价格会直接入库（NULL 也入库）。
     * @return 校验失败返回错误提示，通过返回 null
     */
    private String validateBook(Book book) {
        if (book == null || book.getName() == null || book.getName().trim().isEmpty()) {
            return "新增/修改失败：书名不能为空";
        }
        if (book.getPrice() == null || book.getPrice().compareTo(BigDecimal.ZERO) < 0) {
            return "新增/修改失败：价格必须为不小于0的数字";
        }
        if (book.getSales() == null || book.getSales() < 0) {
            return "新增/修改失败：销量不能为负数";
        }
        if (book.getStock() == null || book.getStock() < 0) {
            return "新增/修改失败：库存不能为负数";
        }
        return null;
    }

    protected void add(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        int pageNo = WebUtils.parseInt(req.getParameter("pageNo"),0);
        pageNo+=1;
        Book book = (Book) WebUtils.copyParamToBean(req.getParameterMap(),new Book());
        // 修复 BUG-M6-01：入库前服务端校验
        String error = validateBook(book);
        if (error != null) {
            req.setAttribute("msg", error);
            req.setAttribute("book", book);
            req.getRequestDispatcher("/pages/manager/book_edit.jsp").forward(req, resp);
            return;
        }
        bookService.addBook(book);
        //req.getRequestDispatcher("/manager/bookServlet?action=list").forward(req,resp);
        resp.sendRedirect(req.getContextPath() + "/manager/bookServlet?action=page&pageNo="+pageNo);
    }


    protected void delete(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
       // 修复 BUG-M6-02：非数字 id 容错（原 Integer.parseInt 直接 500），不存在的 id 给出提示（原静默成功）
       int id = WebUtils.parseInt(req.getParameter("id"), -1);
       int affected = bookService.deleteBookById(id);
       String msg = affected > 0 ? "删除成功" : "删除失败：图书不存在";
       resp.sendRedirect(req.getContextPath() + "/manager/bookServlet?action=page&pageNo="
               + req.getParameter("pageNo") + "&msg=" + URLEncoder.encode(msg, "UTF-8"));
    }


    protected void update(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        Book book = (Book) WebUtils.copyParamToBean(req.getParameterMap(),new Book());
        // 修复 BUG-M6-01：更新同样走服务端校验
        String error = validateBook(book);
        if (error != null) {
            req.setAttribute("msg", error);
            req.setAttribute("book", book);
            req.getRequestDispatcher("/pages/manager/book_edit.jsp").forward(req, resp);
            return;
        }
        bookService.updateBook(book);
        resp.sendRedirect(req.getContextPath() + "/manager/bookServlet?action=page&pageNo="+req.getParameter("pageNo"));
    }

    /**
     *
     * @param req
     * @param resp
     * @throws ServletException
     * @throws IOException
     */
    protected void getBook(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        String id = req.getParameter("id");
        int i = Integer.parseInt(id);
        Book book = bookService.queryBookById(i);
        req.setAttribute("book",book);
        req.getRequestDispatcher("/pages/manager/book_edit.jsp").forward(req,resp);
    }


    protected void list(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        //1、通过BookService查询数据
        List<Book>books = bookService.queryBooks();
        //2、将数据保存在request域中
        req.setAttribute("books",books);
        //3、请求转发到pages/manager/book_manager.jsp
        req.getRequestDispatcher("/pages/manager/book_manager.jsp").forward(req,resp);
    }

    protected void page(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        //1、获取请求的参数pageNo和pageSize
        int pageNo = WebUtils.parseInt(req.getParameter("pageNo"),1);
        int pageSize = WebUtils.parseInt(req.getParameter("pageSize"), Page.PAGE_SIZE);

        //2、调用BookService.page(pageNo,pageSize)方法：返回page对象
        Page<Book> page = bookService.page(pageNo,pageSize);
        page.setUrl("manager/bookServlet?action=page");

        //3、保存Page对象到request域中
        req.setAttribute("page",page);
        //4、请求转发到page/manager/book_manager.jsp页面
        req.getRequestDispatcher("/pages/manager/book_manager.jsp").forward(req,resp);
    }
}
