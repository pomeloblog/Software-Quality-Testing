package com.yj.service;

import com.yj.bean.Book;
import com.yj.bean.Page;

import java.math.BigDecimal;
import java.util.List;

/**
 * @author yj
 * @create 2020-08-24 14:40
 */
public interface BookService {

    public void addBook(Book book);

    public void updateBook(Book book);

    /** 删除图书，返回受影响行数（0=不存在），用于给前端提示。修复 BUG-M6-02。 */
    public int deleteBookById(Integer id);

    public Book queryBookById(Integer id);

    public List<Book> queryBooks();

    Page<Book> page(int pageNo, int pageSize);

    /** 价格为 DECIMAL，区间参数改用 double。修复 BUG-M3-04。 */
    Page<Book> pageByPrice(int pageNo, int pageSize, double min, double max);

    Page<Book> pageByNameOrAuthor(int pageNo, int pageSize, String nameOrAuthor);

    Page<Book> pageOrder();

    public Integer queryTotalBooks();

    BigDecimal queryTotalMoney();
}
