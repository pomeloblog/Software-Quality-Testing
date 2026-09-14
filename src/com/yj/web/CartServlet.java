package com.yj.web;

import com.google.gson.Gson;
import com.yj.bean.Book;
import com.yj.bean.Cart;
import com.yj.bean.CartItem;
import com.yj.service.BookService;
import com.yj.service.impl.BookServiceImpl;
import com.yj.utils.WebUtils;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;


public class CartServlet extends BaseServlet {

    private BookService bookService = new BookServiceImpl();

    protected void ajaxAddItem(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        int id = WebUtils.parseInt(req.getParameter("id"),0);
        Book book = bookService.queryBookById(id);
        //修正：图书不存在时（含 id 为非数字被解析成 0 的情况）原先会解引用 null 抛空指针并返回 500
        if (book == null) {
            resp.setContentType("application/json;charset=UTF-8");
            resp.getWriter().write("{\"totalCount\":0,\"lastName\":\"\",\"msg\":\"图书不存在\"}");
            return;
        }
        CartItem cartItem = new CartItem(book.getId(),book.getName(),1,book.getPrice(),book.getPrice());
        Cart cart = (Cart) req.getSession().getAttribute("cart");
        if(cart==null) {
            cart = new Cart();
            req.getSession().setAttribute("cart",cart);
        }
        cart.addItem(cartItem);
        req.getSession().setAttribute("lastName",cartItem.getName());

        //返回购物车总数量和最后一个商品的名称
        Map<String,Object> resultMap = new HashMap<String,Object>();
        resultMap.put("totalCount",cart.getTotalCount());
        resultMap.put("lastName",cartItem.getName());
        Gson gson = new Gson();
        String resultMapJsonString = gson.toJson(resultMap);
        resp.getWriter().write(resultMapJsonString);

    }
    /**
     * 加入购物车
     * @param req
     * @param resp
     * @throws ServletException
     * @throws IOException
     */
    protected void addItem(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
       int id = WebUtils.parseInt(req.getParameter("id"),0);
       Book book = bookService.queryBookById(id);
        //修正：图书不存在时原先会因 book 为 null 抛空指针
        if (book == null) {
            req.getSession().setAttribute("cartMsg","图书不存在！");
            resp.sendRedirect(referer(req));
            return;
        }
       CartItem cartItem = new CartItem(book.getId(),book.getName(),1,book.getPrice(),book.getPrice());
        Cart cart = (Cart) req.getSession().getAttribute("cart");
        if(cart==null) {
            cart = new Cart();
            req.getSession().setAttribute("cart",cart);
        }
        cart.addItem(cartItem);
        req.getSession().setAttribute("lastName",cartItem.getName());
        resp.sendRedirect(referer(req));
    }

    /**
     * 删除商品项
     * @param req
     * @param resp
     * @throws ServletException
     * @throws IOException
     */
    protected void deleteItem(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        int id = WebUtils.parseInt(req.getParameter("id"),0);
        Cart cart = (Cart) req.getSession().getAttribute("cart");
        if(cart!=null) {
            cart.deleteItem(id);
        }
        //修正：原先把跳转写在 if 内部，购物车为空时不产生任何响应，客户端会停在空白页
        resp.sendRedirect(referer(req));
    }

    /**
     * 清空商品项
     * @param req
     * @param resp
     * @throws ServletException
     * @throws IOException
     */
    protected void clearItem(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
       req.getSession().removeAttribute("cart");
        resp.sendRedirect(referer(req));
    }

    /**
     * 修改商品数量
     * @param req
     * @param resp
     * @throws ServletException
     * @throws IOException
     */
    protected void updateCount(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        int id = WebUtils.parseInt(req.getParameter("id"),0);
        //修正：默认值由 1 改为 0，以便区分「未提供/格式非法」与合法数量；
        //      原先非法输入会被静默改成 1，篡改用户数据且不给出任何提示
        int count = WebUtils.parseInt(req.getParameter("count"),0);
        Cart cart = (Cart) req.getSession().getAttribute("cart");
        if(cart!=null) {
            if (count <= 0) {
                req.getSession().setAttribute("cartMsg","商品数量必须是正整数！");
            } else {
                cart.updateCount(id,count);
            }
        }
        //修正：与 deleteItem 同理，购物车为空时也必须给出响应
        resp.sendRedirect(referer(req));
    }

    /**
     * 修正：Referer 缺失时回退到购物车页，避免跳转目标为 null
     */
    private String referer(HttpServletRequest req) {
        String referer = req.getHeader("Referer");
        return referer != null ? referer : req.getContextPath() + "/pages/cart/cart.jsp";
    }

}
