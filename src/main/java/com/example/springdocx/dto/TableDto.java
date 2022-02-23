package com.example.springdocx.dto;

import io.swagger.models.auth.In;
import lombok.Data;

import java.util.List;

/**
 * @description:
 * @author: dongkunshuai
 * @date: 2022/2/22 14:39
 */
@Data
public class TableDto {

    /**
     * 表格索引
     */
    private Integer index;

    /**
     * 表格标题
     */
    private String title;
    /**
     * 表格存在数据条数
     */
    private Integer count;

    /**
     * 是否计算合计 1. 是 0 否
     */
    private Integer isHj;

    /**
     *  表格项目类型 1 年初+ 年末 2 本年累计
     */
    private Integer type;

    /**
     * 项目列表
     */
    private List<ItemType> itemTypes;


}
