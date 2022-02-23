package com.example.springdocx.util;

import com.example.springdocx.dto.ItemType;
import com.example.springdocx.dto.ItemType1;
import com.example.springdocx.dto.ItemType2;
import com.example.springdocx.dto.TableDto;
import com.example.springdocx.enums.TableHeaderEnum;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @description: word 动态生成表格工具类
 * @author: dongkunshuai
 * @date: 2022/2/22 16:22
 */
@SuppressWarnings("all")
public class DynamicWord {


    /**
     * 动态生成表格工具类
     * @param doc doc 解析word 文档
     * @param tables 生成表格的参数
     * @param key 动态生成表格的位置
     */
    public static void processCreateTable(XWPFDocument doc, List<TableDto> tables, String key){
        try {

            List<XWPFParagraph> paragraphs = doc.getParagraphs();
            if(paragraphs != null && paragraphs.size() > 0){

                for (int p = 0; p < paragraphs.size(); p++) {
                    XWPFParagraph paragraph = paragraphs.get(p);

                    List<XWPFRun> runs = paragraph.getRuns();
                    for (int i = 0; i < runs.size(); i++) {
                        String text = runs.get(i).getText(0);
                        if (text != null) {
                            text = text.trim();
                            if (text.indexOf(key) >= 0) {
                                //从这里开始动态生成表格

                                createTable(doc, paragraph, tables, runs.get(i));

                            }
                        }


                    }


                }



            }



        }catch (Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 数据动态生成表格
     * @param doc
     * @param paragraph
     * @param tables
     * @param run
     */
    private static void createTable(XWPFDocument doc, XWPFParagraph paragraph, List<TableDto> tables, XWPFRun run){


        tables = tables.stream().sorted(Comparator.comparing(TableDto::getIndex).reversed()).collect(Collectors.toList());
        List<TableDto> collect = tables.stream().filter(e -> e.getCount() != 0).collect(Collectors.toList());
        int maxCount = collect.size();
        int count = maxCount;
        //设置新的光标的位置
        XmlCursor newCursor = null;

        for(TableDto tableDto : tables){

            if(tableDto.getCount() != 0){
                //分类型

                if(count == maxCount){
                    run.setText(count + "、" + tableDto.getTitle(), 0);

                    //从段落中获取光标
                    XmlCursor cursor = paragraph.getCTP().newCursor();
                    newCursor = cursor.newCursor();

                    cursor.toNextSibling();
                    //将光标移动到下一个段落
                    insertTableInCursor(doc, cursor, tableDto);
                    count --;
                }else{
                    XWPFParagraph xwpfParagraph = doc.insertNewParagraph(newCursor);
                    XWPFRun r1 = xwpfParagraph.createRun();
                    r1.setFontFamily("宋体");
                    r1.setFontSize(12);
                    r1.setTextPosition(0);
                    r1.setBold(true);
                    r1.addBreak(); // 换行
                    r1.setText(count + "、" + tableDto.getTitle());
                    XmlCursor newCursor1 = xwpfParagraph.getCTP().newCursor();
                    newCursor = newCursor1.newCursor();
                    newCursor1.toNextSibling();
                    insertTableInCursor(doc, newCursor1, tableDto);
                    count -- ;
                }


            }

        }

    }


    /**
     * 生成表格工具
     * @param doc
     * @param cursor
     * @param tableDto
     * @param run
     * @param count
     */
    private static void insertTableInCursor(XWPFDocument doc, XmlCursor cursor, TableDto tableDto){

        XWPFTable table = doc.insertNewTbl(cursor);
        setTableStyle(table);
        if(tableDto.getType() == 1){
            table.addNewCol();
            table.addNewCol();
        }else{
            table.addNewCol();
        }

        //设置有多少行, 项目数 + 表头
        int count = tableDto.getItemTypes().size() + 1;
        if(tableDto.getIsHj() == 1){
            //如果带合计再添加一行
            count += 1;
        }
        for (int i = 0; i < count; i++){
            table.createRow();
        }

        //遍历表格插入数据
        List<XWPFTableRow> rows = table.getRows();
        for (int k = 1; k < rows.size(); k++) {
            List<XWPFTableCell> cells = rows.get(k).getTableCells();
            for (int h = 0; h < cells.size(); h++) {
                XWPFTableCell cell = cells.get(h);
                if(k == 1){
                    createHeader(tableDto, cell, h);
                }else {

                    if(tableDto.getIsHj() == 1){

                        if(k == rows.size() - 1){
                            createStatistics(tableDto, cell, h);
                        }else{
                            createContent(tableDto, cell, k, h);
                        }

                    }else{
                        createContent(tableDto, cell, k, h);
                    }



                }



            }
        }
        table.removeRow(0); //删除第一行
        table.setCellMargins(100,0,100,0);

    }


    public static void createStatistics(TableDto tableDto, XWPFTableCell cell, int index){

        if(tableDto.getType() == 1){
            String content = "合计";
            List<ItemType> itemTypes = tableDto.getItemTypes();

            if(index == 1){
                BigDecimal result = new BigDecimal(0);
                result.setScale(2);
                for (int i = 0; i < itemTypes.size(); i ++){
                    ItemType1 itemType1 = (ItemType1) itemTypes.get(i);
                    if(itemType1.getItemName().contains("减")){
                        if(itemType1.getNcs() != null && !itemType1.getNcs().equals("")){
                            result = result.subtract(new BigDecimal(itemType1.getNcs()));
                        }

                    }else{
                        if(itemType1.getNcs() != null && !itemType1.getNcs().equals("")){
                            result = result.add(new BigDecimal(itemType1.getNcs()));
                        }
                    }
                }
                if(result.compareTo(new BigDecimal(0)) != 0){
                    content = result.toString();
                }else{
                    content = "";
                }

            }else if(index == 2){
                BigDecimal result = new BigDecimal(0);
                result.setScale(2);
                for (int i = 0; i < itemTypes.size(); i ++){
                    ItemType1 itemType1 = (ItemType1) itemTypes.get(i);
                    if(itemType1.getItemName().contains("减")){
                        if(itemType1.getNms() != null && !itemType1.getNms().equals("")){
                            result = result.subtract(new BigDecimal(itemType1.getNms()));
                        }

                    }else{
                        if(itemType1.getNms() != null && !itemType1.getNms().equals("")){
                            result = result.add(new BigDecimal(itemType1.getNms()));
                        }
                    }
                }
                if(result.compareTo(new BigDecimal(0)) != 0){
                    content = result.toString();
                }else{
                    content = "";
                }
            }

            if(tableDto.getType() == 1){
                setCell(cell, content, 2800);
            }else {
                setCell(cell, content, 4200);
            }

        }else{
            String content = "合计";
            List<ItemType> itemTypes = tableDto.getItemTypes();

            if(index == 1){
                BigDecimal result = new BigDecimal(0);
                result.setScale(2);
                for (int i = 0; i < itemTypes.size(); i ++){
                    ItemType2 itemType1 = (ItemType2) itemTypes.get(i);
                    if(itemType1.getItemName().contains("减")){
                        if(itemType1.getTotal() != null && !itemType1.getTotal().equals("")){
                            result = result.subtract(new BigDecimal(itemType1.getTotal()));
                        }

                    }else{
                        if(itemType1.getTotal() != null && !itemType1.getTotal().equals("")){
                            result = result.add(new BigDecimal(itemType1.getTotal()));
                        }
                    }
                }

                if(result.compareTo(new BigDecimal(0)) != 0){
                    content = result.toString();
                }else{
                    content = "";
                }


            }
            if(tableDto.getType() == 1){
                setCell(cell, content, 2800);
            }else {
                setCell(cell, content, 4200);
            }
        }


    }




    public static void createContent(TableDto tableDto, XWPFTableCell cell, int k, int index){

        if(tableDto.getType() == 1){
            List<ItemType> itemTypes = tableDto.getItemTypes();
            ItemType1 itemType = (ItemType1) itemTypes.get(k - 2);
            String content = itemType.getItemName();
            if(index == 0) {
                content = itemType.getItemName();
            }else if(index == 1){
                content = itemType.getNcs();
            }else if(index == 2){
                content = itemType.getNms();
            }
            // 设置水平居中,须要ooxml-schemas包支持
            if(tableDto.getType() == 1){
                setCell(cell, content, 2800);
            }else {
                setCell(cell, content, 4200);
            }

        }else{
            List<ItemType> itemTypes = tableDto.getItemTypes();
            ItemType2 itemType = (ItemType2) itemTypes.get(k - 2);
            String content = itemType.getItemName();
            if(index == 0) {
                content = itemType.getItemName();
            }else if(index == 1){
                content = itemType.getTotal();
            }
            // 设置水平居中,须要ooxml-schemas包支持
            if(tableDto.getType() == 1){
                setCell(cell, content, 2800);
            }else {
                setCell(cell, content, 4200);
            }
        }

    }



    /**
     * 创建表头和表内容
     * @param tableDto
     * @param cell
     * @param index  对应相对应的数据的下标
     */
    public static void  createHeader(TableDto tableDto, XWPFTableCell cell, int index){
        String[] header = TableHeaderEnum.getHeader(tableDto.getType());
        String content = header[index];
        // 设置水平居中,须要ooxml-schemas包支持
        if(tableDto.getType() == 1){
            setCell(cell, content, 2800);
        }else {
            setCell(cell, content, 4200);
        }
    }


    /**
     * 设置表格格式
     * @param table
     */
    private static void setTableStyle(XWPFTable table){

        CTTblPr tblPr =  table.getCTTbl().getTblPr();
        CTTblLayoutType t = tblPr.isSetTblLayout()?tblPr.getTblLayout():tblPr.addNewTblLayout();
        t.setType(STTblLayoutType.FIXED);

        table.setWidth(3000); // 占页面宽度90%




        //设置表格格式
        CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
        CTBorder tBorder = borders.addNewTop();
        String bolderType = "double";
        tBorder.setVal(STBorder.Enum.forString(bolderType));
        tBorder.setSz(new BigInteger("1"));
        tBorder.setColor("000000");

        CTBorder bBorder = borders.addNewBottom();
        bBorder.setVal(STBorder.Enum.forString(bolderType));
        bBorder.setSz(new BigInteger("1"));
        bBorder.setColor("000000");

        CTBorder lBorder=borders.addNewLeft();
        lBorder.setVal(STBorder.Enum.forString("none"));
        lBorder.setSz(new BigInteger("1"));
        lBorder.setColor("3399FF");

        CTBorder rBorder=borders.addNewRight();
        rBorder.setVal(STBorder.Enum.forString("none"));
        rBorder.setSz(new BigInteger("1"));
        rBorder.setColor("F2B11F");


        CTBorder hBorder = borders.addNewInsideH();
        hBorder.setVal(STBorder.Enum.forString("dotted"));
        hBorder.setSz(new BigInteger("1"));
        hBorder.setColor("000000");

        CTBorder vBorder = borders.addNewInsideV();
        vBorder.setVal(STBorder.Enum.forString("dotted"));
        vBorder.setSz(new BigInteger("1"));
        vBorder.setColor("000000");


    }


    /**
     * 设置单元格
     * @param cell
     * @param content
     * @param width
     */
    private static void setCell(XWPFTableCell cell, String content, Integer width){
        // 设置水平居中,须要ooxml-schemas包支持
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);

        cell.setText(content);

        CTTcPr ctTcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth ctTblWidth = ctTcPr.addNewTcW();
        ctTblWidth.setW(BigInteger.valueOf(width));
        ctTblWidth.setType(STTblWidth.DXA);
    }









}
