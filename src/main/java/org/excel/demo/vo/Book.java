package org.excel.demo.vo;

import org.excel.demo.annotation.Excel;

public class Book {


    @Excel(name="필드1", width = 10)
    private String a = "aaaaaaa";

    @Excel(name="필드2", width = 20)
    private static String B = "BOOK";

    @Excel(name="필드3", width = 30)
    private static final String C = "ccccc";

//    @Excel(name="필드4", width = 40)
    private String d = "dddddd";

    @Excel(name="필드5", width = 50)
    private String e = "eeeeee";

    public Book() {
    }

    public Book(String a, String d, String e) {
        this.a = a;
        this.d = d;
        this.e = e;
    }

    private void f() {
        System.out.println("F");
    }

    public void g() {
        System.out.println("g");
    }

    public int h() {
        return 100;
    }
}
