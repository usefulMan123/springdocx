package com.example.springdocx.util;

import org.apache.poi.POIXMLDocument;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.beans.IntrospectionException;
import java.io.FileOutputStream;
import java.lang.reflect.InvocationTargetException;
import java.math.BigInteger;
import java.util.List;

/**
 * @description:
 * @author: dongkunshuai
 * @date: 2022/2/18 14:06
 */

@SuppressWarnings("all")
public class InsertExcelByKeyInWord {

    public static int[] COLUMN_WIDTHS = new int[] {1504,1504,1504,1504,1504,1504};

    //创建一个表格插入到key标记的位置
    public <T>void exportBg(List<String> head,List<T> data,String srcPath,String targetPath,String key) {
        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(POIXMLDocument.openPackage(srcPath));
            int count = 1;
            List<XWPFParagraph> paragraphList = doc.getParagraphs();
            if (paragraphList != null && paragraphList.size() > 0) {
                for (int p = 0; p < paragraphList.size(); p++) {
                    XWPFParagraph paragraph = paragraphList.get(p);
                    List<XWPFRun> runs = paragraph.getRuns();
                    for (int i = 0; i < runs.size(); i++) {
                        String text = runs.get(i).getText(0);

                        if (text != null) {
                            text = text.trim();
                            if (text.indexOf(key) >= 0) {


                                if(count > 1){
                                    XWPFRun xwpfRun = runs.get(i);

                                    paragraph.removeRun(i);
                                    break;
                                }
                                runs.get(i).setText(text.replace(key, count + "、货币资金"), 0);


                                //从段落中获取光标
                                XmlCursor cursor = paragraph.getCTP().newCursor();
                                XmlCursor newCursor = cursor.newCursor();

                                cursor.toNextSibling();
                                //将光标移动到下一个段落
                                createTable(doc, cursor);

                                XWPFParagraph xwpfParagraph = doc.insertNewParagraph(newCursor);
                                XWPFRun r1 = xwpfParagraph.createRun();
                                r1.setFontFamily("宋体");
                                r1.setFontSize(12);
                                r1.setTextPosition(0);
                                r1.setBold(true);
                                r1.addBreak(); // 换行
                                r1.setText("2、呵呵哒测试");
                                XmlCursor newCursor1 = xwpfParagraph.getCTP().newCursor();
                                newCursor1.toNextSibling();
                                createTable(doc, newCursor1);
                                count ++;

                                break ;
                            }
                        }
                    }
                }
            }
            FileOutputStream os = new FileOutputStream(targetPath);
            doc.write(os);
            os.flush();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    /**
     * 复制段落
     *
     * @param sourcePar 原段落
     * @param targetPar
     * @param texts
     */
    private void copyParagraph(XWPFParagraph sourcePar, XWPFParagraph targetPar, String... texts) {

        targetPar.setAlignment(sourcePar.getAlignment());
        targetPar.setVerticalAlignment(sourcePar.getVerticalAlignment());

        // 设置布局
        targetPar.setAlignment(sourcePar.getAlignment());
        targetPar.setVerticalAlignment(sourcePar.getVerticalAlignment());

        if (texts != null && texts.length > 0) {
            String[] arr = texts;
            XWPFRun xwpfRun = sourcePar.getRuns().size() > 0 ? sourcePar.getRuns().get(0) : null;

            for (int i = 0, len = texts.length; i < len; i++) {
                String text = arr[i];
                XWPFRun run = targetPar.createRun();

                run.setText(text);

                run.setFontFamily(xwpfRun.getFontFamily());
                int fontSize = xwpfRun.getFontSize();
                run.setFontSize((fontSize == -1) ? 12 : fontSize);
                run.setBold(xwpfRun.isBold());
                run.setItalic(xwpfRun.isItalic());
            }
        }

    }

    public static void createTable(XWPFDocument doc , XmlCursor cursor){
        XWPFTable table = doc.insertNewTbl(cursor);
        table.addNewCol();
        table.addNewCol();

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

        //设置内容为空
        for (int j = 1; j < 4; j++) {
            table.createRow();
        }




        //遍历表格插入数据
        List<XWPFTableRow> rows = table.getRows();
        for (int k = 1; k < rows.size(); k++) {
            List<XWPFTableCell> cells = rows.get(k).getTableCells();
            for (int h = 0; h < cells.size(); h++) {
                XWPFTableCell cell = cells.get(h);

                // 设置水平居中,须要ooxml-schemas包支持
                CTTc cttc = cell.getCTTc();
                CTTcPr ctPr = cttc.addNewTcPr();
                ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);

                cell.setText("项目");

                CTTcPr ctTcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
                CTTblWidth ctTblWidth = ctTcPr.addNewTcW();
                ctTblWidth.setW(BigInteger.valueOf(2800));

                ctTblWidth.setType(STTblWidth.DXA);
            }
        }
        table.removeRow(0); //删除第一行
        table.setCellMargins(100,0,100,0);
    }


}
