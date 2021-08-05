package com.daxuexiu.excel.entity;

import lombok.Data;

@Data
public class FieldProperty {

    private String name;

    private Integer index;

    private String columnName;

    @Override
    public String toString() {
        return "FieldProperty{" +
                "name='" + name + '\'' +
                ", index=" + index +
                ", columnName='" + columnName + '\'' +
                '}';
    }
}
