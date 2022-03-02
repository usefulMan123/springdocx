package com.example.springdocx;

import com.example.springdocx.util.ExcelReplaceUtil;
import com.example.springdocx.util.StringUtils;
import com.example.springdocx.vo.TableFzVo;
import com.example.springdocx.vo.TableLrVo;
import com.example.springdocx.vo.TableXjVo;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.List;

/**
 * @description:
 * @author: dongkunshuai
 * @date: 2022/2/28 16:38
 */

@SpringBootTest
public class SpringExcelTestApplication {




    @Test
    public void contextLoad2(){

        String srcPath = "D:/fileOpera/template/资产负债表.xls";
        String srcPath2 = "D:/fileOpera/template/利润表.xls";
        String srcPath3= "D:/fileOpera/template/现金流量表.xls";

        String srcPath4 = "D:/fileOpera/template/zcfz.xlsx";
        try {
            List<TableFzVo> table = (List<TableFzVo>) ExcelReplaceUtil.getTable(1, srcPath, null);
            List<TableLrVo> table2 = (List<TableLrVo>) ExcelReplaceUtil.getTable(2, srcPath2, null);
            List<TableXjVo> table1 = (List<TableXjVo>) ExcelReplaceUtil.getTable(3, srcPath3, null);


            ExcelReplaceUtil.replaceExcel(srcPath4, table, table2, table1, null);
            System.out.println("文件生成成功");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }



}
