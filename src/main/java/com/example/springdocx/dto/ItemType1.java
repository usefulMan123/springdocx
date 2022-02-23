package com.example.springdocx.dto;

import lombok.Data;

/**
 * @description: 项目名+年初+年末组合
 * @author: dongkunshuai
 * @date: 2022/2/22 14:43
 */

@Data
public class ItemType1 extends ItemType{


    /**
     * 年初数
     */
    private String ncs;

    /**
     * 年末数
     */
    private String nms;

}
