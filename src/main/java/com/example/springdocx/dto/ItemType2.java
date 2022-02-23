package com.example.springdocx.dto;

import lombok.Data;

/**
 * @description: 项目类型2项目名 + 本年累计
 * @author: dongkunshuai
 * @date: 2022/2/22 14:49
 */

@Data
public class ItemType2 extends ItemType{


    /**
     * 本年累计
     */
    private String total;


}
