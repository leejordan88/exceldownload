package org.excel.demo.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.excel.demo.annotation.Excel;
import org.springframework.stereotype.Service;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Service
public class ExcelService {

    private static final String SHEET = "sheet";
    private static final int SHEET_SIZE = 65000;
    private static final int FIXEL = 500;

    public SXSSFWorkbook makeSimpleExcelWorkbook(List<?> list, boolean setHeaderStyle) throws IllegalAccessException {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        List<String> headerNameList = new ArrayList<>();
        List<Integer> headerWidthList = new ArrayList<>();

        Class<?> cls = list.get(0).getClass();
        Arrays.stream(cls.getDeclaredFields()).forEach(f -> {
            f.setAccessible(true);
            Excel annotation = f.getAnnotation(Excel.class);
            if (annotation != null) {
                headerNameList.add(annotation.name());
                headerWidthList.add(annotation.width());
            }
        });

        // 시트 정보
        int sheetLength = (list.size() < SHEET_SIZE) ? 1 : list.size() / SHEET_SIZE;
        sheetLength = (sheetLength != 1 && (sheetLength % SHEET_SIZE == 0)) ? sheetLength : sheetLength + 1;
        int shhetIndex = 0;
        int rowIndex = 0;
        int cellIndex = 0;
        int bodyIndex = 0;

        for (int i = shhetIndex; i < sheetLength; i++) {
            //첫 시트 생성
            SXSSFSheet sheet = workbook.createSheet(SHEET + i);

            //첫 로우 생성
            Row headerRow = sheet.createRow(rowIndex++);

            //해더 정보
            for (String name : headerNameList) {
                sheet.setColumnWidth(cellIndex, name.length() * FIXEL);

                //Custom header setting
                if (setHeaderStyle) {
                    sheet.setColumnWidth(cellIndex, headerWidthList.get(cellIndex) * FIXEL);
                    //TODO 컬러 셋팅
                }
                // 해당 행의 첫번째 열 셀 생성
                Cell headerCell = headerRow.createCell(cellIndex);
                headerCell.setCellValue(name);
                cellIndex++;
            }
            
            //바디 정보
            int lastSheetCount = (list.size() < SHEET_SIZE) ? list.size()
                    : (bodyIndex + SHEET_SIZE > list.size()) ? list.size() : bodyIndex + SHEET_SIZE;
            for (int j = bodyIndex; j < lastSheetCount; j++) {

                if (j == lastSheetCount - 1) {
                    bodyIndex += SHEET_SIZE;
                }

                Object o = list.get(j);
                Class<?> bodyClass = o.getClass();
                Row bodyRow = sheet.createRow(rowIndex++);

                cellIndex = 0;
                Field[] declaredFields = bodyClass.getDeclaredFields();
                for (Field f : declaredFields) {
                    f.setAccessible(true);
                    Cell bodyCell = bodyCell = bodyRow.createCell(cellIndex++);
                    bodyCell.setCellValue(String.valueOf(f.get(o)));
                }
            }
            shhetIndex++;
            rowIndex = 0;
            cellIndex = 0;
        }
        return workbook;
    }

    public SXSSFWorkbook excelFileDownloadProcess(List<?> list, boolean setHeaderStyle) throws IllegalAccessException {
        return this.makeSimpleExcelWorkbook(list, setHeaderStyle);
    }

}
