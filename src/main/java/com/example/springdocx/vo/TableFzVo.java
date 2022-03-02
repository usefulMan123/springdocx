package com.example.springdocx.vo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * @description:
 * @author: dongkunshuai
 * @date: 2022/2/28 17:46
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class TableFzVo implements Serializable {

    //项目名字
    private String itemName;

    //获取年初数
    private String ncs;

    //获取年末数
    private String nms;

}
