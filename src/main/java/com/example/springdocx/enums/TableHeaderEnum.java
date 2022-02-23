package com.example.springdocx.enums;

/**
 * @description: 设置表格对应表头-- 快写完了发现还有表头没有设计进去， 只能这么搞了
 * @author: dongkunshuai
 * @date: 2022/2/22 17:22
 */
@SuppressWarnings("all")
public enum  TableHeaderEnum {
    TABLETYPE1(1, new String[]{"项目", "年初数", "年末数"}),
    TABLETYPE2(2, new String[]{"项目", "本年累计数"})

    ;


    public static String[] getHeader(Integer type){

        for(TableHeaderEnum tableHeaderEnum : TableHeaderEnum.values()){
            if(tableHeaderEnum.getType() == type){
                return tableHeaderEnum.getHeader();
            }
        }

        return null;

    }



    private Integer type;

    private String[] header;


    TableHeaderEnum() {
    }

    TableHeaderEnum(Integer type, String[] header) {
        this.type = type;
        this.header = header;
    }

    public String[] getHeader() {
        return header;
    }

    public void setHeader(String[] header) {
        this.header = header;
    }

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
    }
}
