package com.example.springdocx.vo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * @description: 利润表
 * @author: dongkunshuai
 * @date: 2022/3/1 9:31
 */

@Data
@AllArgsConstructor
@NoArgsConstructor
public class TableLrVo  implements Serializable {

    //项目
    private String itemName;

    //本季累计金额
    private String bjMoney;

    //本年累计金额
    private String bnMoney;

}
