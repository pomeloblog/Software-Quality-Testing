package com.yj.service.impl;

import com.yj.bean.Book;
import com.yj.bean.Page;
import com.yj.dao.BookDao;
import com.yj.dao.impl.BookDaoImpl;
import com.yj.service.BookService;

import java.math.BigDecimal;
import java.util.List;

/**
 * @author yj
 * @create 2020-08-24 14:50
 */
public class BookServiceImpl implements BookService {


    private BookDao bookDao = new BookDaoImpl();

    @Override
    public void addBook(Book book) {
        bookDao.addBook(book);
    }

    @Override
    public void updateBook(Book book) {
        bookDao.updateBook(book);
    }

    @Override
    public int deleteBookById(Integer id) {
        // 修复 BUG-M6-02：返回受影响行数，供 Servlet 判断"图书不存在"并提示
        return bookDao.deleteBookById(id);
    }

    @Override
    public Book queryBookById(Integer id) {
        return bookDao.queryBookById(id);
    }

    @Override
    public List<Book> queryBooks() {
        return bookDao.queryBooks();
    }

    @Override
    public Page<Book> page(int pageNo, int pageSize) {
        Page<Book> page = new Page<Book>();

        // 修复 BUG-M3-01：pageSize 非法（<1）时容错为默认值，避免除零/负数 LIMIT 导致 500
        if (pageSize < 1) {
            pageSize = Page.PAGE_SIZE;
        }

        //设置每页记录数
        page.setPageSize(pageSize);

        //设置总记录数
        Integer pageTotalCount = bookDao.queryForPageTotalCount();
        page.setPageTotalCount(pageTotalCount);

        //求总页码
        Integer pageTotal = pageTotalCount / pageSize;
        if(pageTotalCount % pageSize >0) {
            pageTotal+=1;
        }
        //设置当前页
        if(pageNo>pageTotal) {
            pageNo = pageTotal;
        }
        if(pageNo<1) {
            pageNo = 1;
        }
        page.setPageNo(pageNo);
        //设置总页码
        page.setPageTotal(pageTotal);


        int begin = (page.getPageNo() -1)*pageSize;
        List<Book> items = bookDao.queryForPageItems(begin,pageSize);
        page.setItems(items);

         return page;
    }

    @Override
    public Page<Book> pageByPrice(int pageNo, int pageSize, double min, double max) {
        Page<Book> page = new Page<Book>();

        // 修复 BUG-M3-01：pageSize 非法（<1）时容错为默认值
        if (pageSize < 1) {
            pageSize = Page.PAGE_SIZE;
        }

        //设置每页记录数
        page.setPageSize(pageSize);

        //设置总记录数
        Integer pageTotalCount = bookDao.queryForPageTotalCountByPrice(min,max);
        page.setPageTotalCount(pageTotalCount);

        //求总页码
        Integer pageTotal = pageTotalCount / pageSize;
        if(pageTotalCount % pageSize >0) {
            pageTotal+=1;
        }
        //设置当前页
        if(pageNo>pageTotal) {
            pageNo = pageTotal;
        }
        if(pageNo<1) {
            pageNo = 1;
        }
        page.setPageNo(pageNo);
        //设置总页码
        page.setPageTotal(pageTotal);


        int begin = (page.getPageNo() -1)*pageSize;
        List<Book> items = bookDao.queryForPageItemsByPrice(begin,pageSize,min,max);
        page.setItems(items);

        return page;
    }

    @Override
    public Page<Book> pageByNameOrAuthor(int pageNo, int pageSize, String nameorauthor) {
        Page<Book> page = new Page<Book>();

        // 修复 BUG-M3-01：pageSize 非法（<1）时容错为默认值
        if (pageSize < 1) {
            pageSize = Page.PAGE_SIZE;
        }

        //设置每页记录数
        page.setPageSize(pageSize);

        //设置总记录数
        Integer pageTotalCount = bookDao.queryForPageTotalCountByNameOrAuthor(nameorauthor);
        page.setPageTotalCount(pageTotalCount);

        //求总页码
        Integer pageTotal = pageTotalCount / pageSize;
        if(pageTotalCount % pageSize >0) {
            pageTotal+=1;
        }
        //设置当前页
        if(pageNo>pageTotal) {
            pageNo = pageTotal;
        }
        if(pageNo<1) {
            pageNo = 1;
        }
        page.setPageNo(pageNo);
        //设置总页码
        page.setPageTotal(pageTotal);


        int begin = (page.getPageNo() -1)*pageSize;
        List<Book> items = bookDao.queryForPageItemsByNameOrAuthor(begin,pageSize,nameorauthor);
        page.setItems(items);

        return page;
    }

    @Override
    public Page<Book> pageOrder() {
        Page<Book> page = new Page<Book>();
        List<Book> items = bookDao.queryForPageItemsOrder();
        page.setItems(items);

        return page;
    }

    @Override
    public Integer queryTotalBooks() {
        return bookDao.queryBooknums();
    }

    @Override
    public BigDecimal queryTotalMoney() {
        return bookDao.queryTotalMoney();
    }


}
