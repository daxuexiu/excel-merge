package com.daxuexiu.excel.entity;

import com.daxuexiu.excel.annotation.Excel;
import lombok.Data;

@Data
public class User {

     @Excel(title = "编号",index = 0,rowIndex=3)
     private String id;

     @Excel(title = "姓名",index = 1,rowIndex=3)
     private String name;

     @Excel(title = "年龄",index = 2,rowIndex=3)
     private String age;
}
