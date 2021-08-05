package com.daxuexiu.excel.entity;

import com.daxuexiu.excel.annotation.Excel;
import com.daxuexiu.excel.annotation.MergeExcel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Table {

    @MergeExcel(value = "行合并", rowspan = 1, colIndex = 0)
    @Excel(title = "用户名", index = 0, rowIndex = 3)
    private String username;

    @Excel(title = "姓名", index = 1, rowIndex = 3)
    private String name;

    @MergeExcel(value = "列合并",rowspan = 1,colIndex = 2, colspan = 1)
    @Excel(title = "性别", index = 2, rowIndex = 3)
    private String sex;

    @Excel(title = "年龄", index = 3, rowIndex = 3)
    private Integer age;

    @Excel(title = "时间", index = 4, rowIndex = 3)
    private Date date;
}
