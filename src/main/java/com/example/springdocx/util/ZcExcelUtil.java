package com.example.springdocx.util;

import com.example.springdocx.dto.ItemType;
import com.example.springdocx.dto.ItemType1;
import com.example.springdocx.dto.ItemType2;
import com.example.springdocx.dto.TableDto;
import com.example.springdocx.enums.ItemEnum;
import com.example.springdocx.enums.TableEnum;
import com.google.common.collect.Tables;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.*;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.stream.Collectors;

/**
 * @description: 资产清算Excel数据读取
 * @author: dongkunshuai
 * @date: 2021/12/28 14:08
 */

public class ZcExcelUtil {


    /**
     * 从Excel中读取模板替换元素集合
     * @param filePath 上传Excel 路径
     * @param templateData 替换的模板集合
     * @return 替换的模板集合
     * @throws Exception
     */
    public static List<TableDto> getTemplateFromExcel(String filePath, Map<String, String> templateData, MultipartFile multipartFile) throws Exception {
        File file= null;
        if (filePath != null) {
            file = new File(filePath);

            if(!file.exists()){
                throw new Exception("Excel 文件地址错误不存在");
            }
        }
        // 存放读取的所有结果
        List<TableDto> tables = new CopyOnWriteArrayList<>();
        Workbook workBook = null;
        try {

            if(file != null) {
                if (file.getName().endsWith("xlsx")) {
                    workBook = new XSSFWorkbook(file);
                } else if (file.getName().endsWith("xls")) {
                    workBook = new HSSFWorkbook(new FileInputStream(file));
                }
            }else{
                if (multipartFile.getOriginalFilename().endsWith("xlsx")) {
                    workBook = new XSSFWorkbook(multipartFile.getInputStream());
                } else if (multipartFile.getOriginalFilename().endsWith("xls")) {
                    workBook = new HSSFWorkbook(multipartFile.getInputStream());
                }
            }

            // 暂时写死， 第三页文档不需要读取数据
            for (int i = 0; i < workBook.getNumberOfSheets() - 1; i++ ){

                // 获取第i页数据
                Sheet sheetAt = workBook.getSheetAt(i);
                //存放第一页的结果
                List<Map<String, String>> sheetList = new ArrayList<>();
                //记录表头的位置
                Map<Integer, String> titleMap = new HashMap<>(sheetAt.getPhysicalNumberOfRows());

                for(int j =4 ; j < sheetAt.getPhysicalNumberOfRows(); j++){
                    //获取第j行数据
                    Row row = sheetAt.getRow(j);
                    if(Objects.isNull(row)){
                        //如果整行为空， 则跳过
                        continue;
                    }

                    for (int k = 0; k < row.getLastCellNum(); k++) {

                        Object cellValue = getCellValue(row.getCell(k));
                        //判断数据是否为空
                        if(cellValue != null){
                            String newCellValue = StringUtils.trim(cellValue.toString());

                            processItemContent(tables, newCellValue, row, k);
                        }

                    }


                }

            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } finally {
            // 关闭资源
            try {
                workBook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return tables;
    }

    /**
     * 处理项目信息并封装为表格
     * @param tables
     * @param cellVaule
     * @param row
     * @param index
     */
    private static void processItemContent(List<TableDto> tables, String cellVaule,  Row row, int index){

        TableDto table = null;
        // 通过列名获取项目数据
        ItemEnum item = ItemEnum.getItem(cellVaule);
        List<ItemType> itemTypes = null;
        if(item != null){
            //获取表格索引
            Integer tableIndex = item.getIndex();


            List<TableDto> collect = tables.stream().filter(e -> tableIndex.equals(e.getIndex())).collect(Collectors.toList());

            if(collect != null && collect.size() != 0){
                table = collect.get(0);
            }else {
                table = new TableDto();
                table.setIsHj(TableEnum.getIsHj(tableIndex));
                table.setTitle(TableEnum.getTitle(tableIndex));
                table.setType(item.getType());
                table.setItemTypes(new CopyOnWriteArrayList<>());
                table.setIndex(tableIndex);
                //默认有0条有数据的条数
                table.setCount(0);
                tables.add(table);
            }

            itemTypes = table.getItemTypes();
            delItem(itemTypes, table, item, cellVaule, row, index);

        }

    }

    private static void delItem(List<ItemType> itemTypes, TableDto table, ItemEnum item,  String cellValue, Row row, int index){


        //区分表格类型
        if(table.getType() == 1) {
            //项目名 + 年初 + 年末
            ItemType1 itemType1 = new ItemType1();
            itemType1.setItemName(item.getItemName());

            //取数
            Object fistE = getCellValue(row.getCell(index + 2));
            Object secondE = getCellValue(row.getCell(index + 3));

            if(!StringUtils.isEmpty(fistE.toString()) || !StringUtils.isEmpty(secondE.toString())){
                table.setCount(table.getCount() + 1);
                itemType1.setNcs(fistE.toString());
                itemType1.setNms(secondE.toString());
            }
            itemTypes.add(itemType1);
        }else if (table.getType() == 2) {


            if(!"未分配利润".equals(cellValue)){

                ItemType2 itemType2 = new ItemType2();
                itemType2.setItemName(item.getItemName());
                Object fistE = getCellValue(row.getCell(index + 2));
                if(!StringUtils.isEmpty(fistE.toString())){
                    itemType2.setTotal(fistE.toString());
                    table.setCount(table.getCount() + 1);
                }

                itemTypes.add(itemType2);


            }else{
                //处理特殊情况
                //取数
                Object fistE = getCellValue(row.getCell(index + 2));
                Object secondE = getCellValue(row.getCell(index + 3));
                if(!StringUtils.isEmpty(fistE.toString())){
                    table.setCount(table.getCount() + 1);
                    ItemType2 itemType2 = new ItemType2();
                    itemType2.setItemName("加：年初未分配利润");
                    itemType2.setTotal(fistE.toString());
                    itemTypes.add(itemType2);
                }
                if(!StringUtils.isEmpty(secondE.toString())){
                    Integer count = table.getCount() == null ? 0 : table.getCount();
                    table.setCount(count + 1);
                    ItemType2 itemType2 = new ItemType2();
                    itemType2.setItemName("年末未分配利润");
                    itemType2.setTotal(secondE.toString());
                    itemTypes.add(itemType2);
                }
            }



        }


    }





    /**
     * 根据单元格的类型获取相应的值
     *
     * @param cell
     * @return
     */
    private static Object getCellValue(Cell cell) {
        Object cellValue;
        if (Objects.nonNull(cell) && cell.getCellTypeEnum() == CellType.NUMERIC) {
            // 数值型
            // poi读取整数会自动转成小数，这里对整数进行还原，小数不会做处理
            double numericCellValue = cell.getNumericCellValue();
            DecimalFormat decimalFormat = new DecimalFormat("0.00");
            String geihua = decimalFormat.format(numericCellValue);
           // System.out.println(geihua);
            /*long longValue = Math.round(cell.getNumericCellValue());
            if (Double.parseDouble(longValue + ".00") == cell.getNumericCellValue()) {
                cellValue = longValue;
            } else {
                cellValue = cell.getNumericCellValue();
            }*/
            cellValue = geihua;
        } else if (Objects.nonNull(cell) && cell.getCellTypeEnum() == CellType.FORMULA) {
            // 公式型
            // 公式计算的值不会转成小数，这里数值获取失败后会获取字符
            try {
                double numericCellValue = cell.getNumericCellValue();
                DecimalFormat format = new DecimalFormat("0.00");
                String format1 = format.format(numericCellValue);

                cellValue = format1;
                /*if(cellValue.equals("0.00")){
                    //公式计算的出的结果是0 就把结果置为空
                    cellValue = "";
                }*/
                //System.out.println(cellValue);
            } catch (Exception e) {
                cellValue = cell.getStringCellValue();
            }
        } else {
            // 其他类型不作处理
            cellValue = cell;
        }
        return cellValue;
    }

    /**
     * 去除Excel手误添加的空格
     *
     * @param str
     * @return
     */
    public static String stringTrim(String str) {
        return str.replaceAll("[\\u00A0]+", "").trim();
    }


    /**
     * 获取指定工作簿，指定行，指定列的数据
     * @param row
     * @param cell
     * @param sheet
     * @return
     */
    public static Object getData(MultipartFile file, Integer row, Integer cell, Integer sheet) throws IOException {
        Workbook workbook = null;
        String name = file.getOriginalFilename();
        System.out.println(name);
        if (file.getOriginalFilename().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else if (file.getOriginalFilename().endsWith("xls")) {
            workbook = new HSSFWorkbook(file.getInputStream());
        }

        Sheet sheetAt = workbook.getSheetAt(sheet);
        Row row1 = sheetAt.getRow(row);
        Object cellValue = getCellValue(row1.getCell(cell));
        return cellValue;
    }

}
