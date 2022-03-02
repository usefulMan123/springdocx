package com.example.springdocx;

import com.example.springdocx.dto.Person;
import com.example.springdocx.dto.TableDto;
import com.example.springdocx.util.DynamicWord;
import com.example.springdocx.util.InsertExcelByKeyInWord;
import com.example.springdocx.util.ZcExcelUtil;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

@SpringBootTest
class SpringdocxApplicationTests {

    @Test
    void contextLoads() {

        System.out.println("hello world");
        InsertExcelByKeyInWord insertExcelByKeyInWord =new InsertExcelByKeyInWord();
        List<Person> dtos=new ArrayList<>();
        for (int i = 0; i <20 ; i++) {
            Person dto=new Person("张三",22,"男");
            dtos.add(dto);
        }
        List<String> head=new ArrayList<>();
        head.add("姓名");
        head.add("年龄");
        head.add("性别");
        String srcPath = "D://fileOpera/template/key.docx";
        String targetPath = "D://fileOpera/template/keyx.docx";
        String key = "${key}";// 在文档中需要替换插入表格的位置
        insertExcelByKeyInWord.exportBg(head,dtos,srcPath,targetPath,key);
        System.out.println("生成结束");
    }

    @Test
    void contextLoads2() {

        String srcPath = "D://fileOpera/template/zcfz.xls";
        try {
            List<TableDto> templateFromExcel = ZcExcelUtil.getTemplateFromExcel(srcPath, null, null);


            System.out.println(templateFromExcel);
            String srcPath2 = "D://fileOpera/template/key.docx";
            XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(srcPath2));
            String key = "${key}";// 在文档中需要替换插入表格的位置
            DynamicWord.processCreateTable(doc, templateFromExcel, key);

            String targetPath = "D://fileOpera/template/keyx.docx";
            FileOutputStream os = new FileOutputStream(targetPath);
            doc.write(os);
            os.flush();
            os.close();


        } catch (Exception e) {
            e.printStackTrace();
        }


    }

}
