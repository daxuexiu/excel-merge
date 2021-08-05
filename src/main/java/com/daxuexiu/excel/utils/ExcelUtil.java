package com.daxuexiu.excel.utils;

import com.daxuexiu.excel.annotation.Excel;
import com.daxuexiu.excel.annotation.MergeExcel;
import com.daxuexiu.excel.entity.FieldProperty;
import com.daxuexiu.excel.entity.Table;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.StringUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;

public class ExcelUtil {

    public static<T> void exportExcel(List<T> list, Class<T> clazz) throws RuntimeException{
        if(clazz == null){
            throw new RuntimeException("类型不能为空");
        }
//        try {
//            T obj = clazz.newInstance();
//        } catch (InstantiationException e) {
//            e.printStackTrace();
//        } catch (IllegalAccessException e) {
//            e.printStackTrace();
//        }

        MergeExcel annotation = clazz.getAnnotation(MergeExcel.class);

        Field[] fields = clazz.getDeclaredFields();

        List<FieldProperty> fieldProperties = new ArrayList<>();
        for(Field field : fields) {
            Excel excel = field.getAnnotation(Excel.class);
            FieldProperty fieldProperty = new FieldProperty();
            fieldProperty.setName(field.getName());
            fieldProperty.setIndex(excel.index());
            fieldProperty.setColumnName(excel.title());

            fieldProperties.add(fieldProperty);
            System.out.println(excel.title() + "----" + excel.index());
        }

        for (FieldProperty fieldProperty : fieldProperties){
            System.out.println(fieldProperty.toString());
        }
    }


    public static void  createTitle(List<FieldProperty> list, HSSFSheet sheet){
        HSSFRow row=sheet.createRow( 1 );
        for (int i = 0; i < list.size(); i++){
            FieldProperty fieldProperty = list.get(i);

            HSSFCell ce = row.createCell(i);
            ce.setCellValue(fieldProperty.getColumnName());
//            ce.setCellStyle(createCellStyle(wb, (short) 10, true, true));
        }
    }

    public static<T> List<FieldProperty> getTitleFiled(Class<T> clazz){
        Field[] fields = clazz.getDeclaredFields();
        List<FieldProperty> fieldProperties = new ArrayList<>();
        for(Field field : fields) {
            Excel excel = field.getAnnotation(Excel.class);
            FieldProperty fieldProperty = new FieldProperty();
            fieldProperty.setName(field.getName());
            fieldProperty.setIndex(excel.index());
            fieldProperty.setColumnName(excel.title());

            fieldProperties.add(fieldProperty);
        }
        return fieldProperties;
    }

    /**
     * 创建Sheet
     * @param workbook
     * @return
     */
    private static HSSFSheet createSheet(HSSFWorkbook workbook, String sheetName) {
        if(StringUtils.isEmpty(sheetName)){
            return workbook.createSheet("Sheet1");
        }else{
            return workbook.createSheet(sheetName);
        }
    }

    /**
     * 创建HSSFWorkbook
     * @return
     */
    private static HSSFWorkbook createWorkBook(){
        return new HSSFWorkbook();
    }

    /**
     * 设置表格的列头部
     * @param sheet
     * @param fields
     * @param style
     */
    private static void setColumnTitle(HSSFSheet sheet, Field[] fields, HSSFCellStyle style) {
        int nextRow = sheet.getLastRowNum() + 1;
        for (Field field : fields) {
            field.setAccessible(true);
            if (field.isAnnotationPresent(Excel.class)) {
                Excel excelColumn = field.getDeclaredAnnotation(Excel.class);
                sheet.setColumnWidth(excelColumn.index(), 15 * 256);
                HSSFRow row = sheet.getRow(nextRow);
                if (Objects.isNull(row)) {
                    row = sheet.createRow(nextRow);
                }
                HSSFCell cell = row.createCell(excelColumn.index());
                cell.setCellValue(excelColumn.title());
                cell.setCellStyle(style);
            }
        }
    }



    private static HSSFCellStyle setHssFCellStyle(HSSFWorkbook workbook){
        HSSFCellStyle style = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static void setMergeExcel(HSSFSheet sheet, Field[] fields, HSSFCellStyle style) {
        int nextRow = sheet.getLastRowNum() + 1;
        for (Field field : fields) {
            field.setAccessible(true);
            if (field.isAnnotationPresent(MergeExcel.class)) {
                MergeExcel template = field.getDeclaredAnnotation(MergeExcel.class);
                // 合并行的数量
                int rowspan = template.rowspan();
                nextRow = rowspan;
                CellRangeAddress region = new CellRangeAddress(nextRow, nextRow + template.rowspan(), template.colIndex(), template.colIndex() + template.colspan());
                HSSFRow row = sheet.getRow(nextRow);
                if (Objects.isNull(row)) {
                    row = sheet.createRow(nextRow);
                }
                sheet.addMergedRegion(region);
                HSSFCell cell = row.createCell(template.colIndex());
                cell.setCellValue(template.value());
                cell.setCellStyle(style);
                HSSFRow lastRow = sheet.getRow(nextRow + template.rowspan());
                if (Objects.isNull(lastRow)) {
                    sheet.createRow(nextRow + template.rowspan());
                }
            }
        }
    }

    /**
     * 获取需要导出字段数量
     * @param clazz
     * @param <T>
     * @return
     */
    private static <T> Field[] getFields(Class<T> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        if (fields == null || fields.length == 0) {
            throw new RuntimeException("clazz：" + clazz.getCanonicalName() + ",实体空异常！");
        }
        for (Field field : fields) {
            if (!field.isAnnotationPresent(Excel.class)) {
                throw new RuntimeException("clazz：" + clazz.getCanonicalName() + ", 实体空Excel注解异常！");
            }
        }
        return fields;
    }


    /**
     * 设置值
     * @param workbook
     * @param sheet
     * @param data
     * @param fields
     * @param <T>
     */
    private static <T> void setData(HSSFWorkbook workbook, HSSFSheet sheet, List<T> data, Field[] fields,HSSFCellStyle style){
        try {
            int lastRow = sheet.getLastRowNum();
            for (int i = 0; i < data.size(); i++) {
                HSSFRow row = sheet.createRow(lastRow + i + 1);
                for (Field field : fields) {
                    field.setAccessible(true);
                    if (field.isAnnotationPresent(Excel.class)) {
                        Excel excelColumn = field.getAnnotation(Excel.class);
                        Object value = field.get(data.get(i));
                        if (Objects.isNull(value)) {
                            continue;
                        }
                        HSSFCell cell = row.createCell(excelColumn.index());
                        cell.setCellValue(value + "");
                        cell.setCellStyle(style);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static void main(String[] args) {
        HSSFWorkbook workBook = createWorkBook();
        HSSFSheet sheet = createSheet(workBook, "测试");

        Field[] fields = getFields(Table.class);
        HSSFCellStyle hssfCellStyle = setHssFCellStyle(workBook);
        setColumnTitle(sheet,fields,hssfCellStyle);
        List<Table> list = new ArrayList<>();
        list.add(new Table("123","1111","男",20,new Date()));
        list.add(new Table("456","2222","女",21,new Date()));
        list.add(new Table("789","33333","女",22,new Date()));
        list.add(new Table("1110","44444","男",23,new Date()));

        setData(workBook,sheet,list,fields,hssfCellStyle);
        setMergeExcel(sheet,fields,hssfCellStyle);


        //输出Excel文件
        try {
            FileOutputStream output = new FileOutputStream("/Users/yueylong/yueylong/code/excel-merge/src/main/resources/students.xls");
            workBook.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
