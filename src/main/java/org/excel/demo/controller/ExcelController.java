package org.excel.demo.controller;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.excel.demo.service.ExcelService;
import org.excel.demo.vo.Book;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import java.util.ArrayList;
import java.util.List;

@Controller
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @RequestMapping(value = "/downloadExcelFile", method = RequestMethod.GET)
    public String downloadExcelFile(Model model) throws IllegalAccessException {
        Book book = new Book();
        List<Book> list = new ArrayList<>();
        for (int i = 0; i < 120000; i++) {
            list.add(book);
        }
        SXSSFWorkbook workbook = excelService.excelFileDownloadProcess(list, true);
        model.addAttribute("workbook", workbook);
        model.addAttribute("workbookName", "테스트");
        return "excelView";
    }

}
