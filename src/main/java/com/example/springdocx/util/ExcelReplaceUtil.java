package com.example.springdocx.util;

import com.example.springdocx.dto.TableDto;
import com.example.springdocx.vo.TableFzVo;
import com.example.springdocx.vo.TableLrVo;
import com.example.springdocx.vo.TableXjVo;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.util.*;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.stream.Collectors;


/**
 * @description: 读取Excel 数据并封装， 及填充Excel的数据
 * @author: dongkunshuai
 * @date: 2022/2/28 16:35
 */

public class ExcelReplaceUtil {


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
     * 获取从负债表，利润表或现金流量表里获取数据的 方法
     * @param type
     * @param filePath
     * @param multipartFile
     * @return
     */
    public static Object getTable(Integer type, String filePath, MultipartFile multipartFile) throws Exception {


        List<Object> list = new ArrayList<>();


        File file = null;
        if (filePath != null) {
            file = new File(filePath);

            if (!file.exists()) {
                throw new Exception("Excel 文件地址错误不存在");
            }
        }
        // 存放读取的所有结果
        List<TableDto> tables = new CopyOnWriteArrayList<>();
        Workbook workBook = null;
        try {

            if (file != null) {
                if (file.getName().endsWith("xlsx")) {
                    workBook = new XSSFWorkbook(file);
                } else if (file.getName().endsWith("xls")) {
                    workBook = new HSSFWorkbook(new FileInputStream(file));
                }
            } else {
                if (multipartFile.getOriginalFilename().endsWith("xlsx")) {
                    workBook = new XSSFWorkbook(multipartFile.getInputStream());
                } else if (multipartFile.getOriginalFilename().endsWith("xls")) {
                    workBook = new HSSFWorkbook(multipartFile.getInputStream());
                }
            }


            //记录行次所在下标
            List<Integer> hcIndex = new ArrayList<>();



            // 暂时写死， 第三页文档不需要读取数据
            for (int i = 0; i < workBook.getNumberOfSheets(); i++ ){

                // 获取第i页数据
                Sheet sheetAt = workBook.getSheetAt(i);

                for(int j =1 ; j < sheetAt.getPhysicalNumberOfRows(); j++){
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
                            String newCellValue = org.springframework.util.StringUtils.trimAllWhitespace(cellValue.toString());
                            if(newCellValue.equals("行次")){
                                hcIndex.add(k);
                            }else if(!newCellValue.equals("行次") && !newCellValue.equals("") && hcIndex.contains(k) && type == 1) {
                                TableFzVo tableFzVo = new TableFzVo();
                                Object tableName = getCellValue(row.getCell(k - 1));
                                Object ncs = getCellValue(row.getCell(k + 1));
                                Object nms = getCellValue(row.getCell(k + 2));
                                tableFzVo.setItemName(org.springframework.util.StringUtils.trimAllWhitespace(tableName.toString()));
                                tableFzVo.setNcs(org.springframework.util.StringUtils.trimAllWhitespace(ncs.toString().trim()));
                                tableFzVo.setNms(org.springframework.util.StringUtils.trimAllWhitespace(nms.toString().trim()));
                                list.add(tableFzVo);
                            }else if(!newCellValue.equals("行次") && !newCellValue.equals("") && hcIndex.contains(k) && type == 2){
                                TableLrVo tableLrVo = new TableLrVo();
                                Object tableName = getCellValue(row.getCell(k - 1));
                                Object bjlj = getCellValue(row.getCell(k + 1));
                                Object bnlj = getCellValue(row.getCell(k + 2));
                                tableLrVo.setItemName(org.springframework.util.StringUtils.trimAllWhitespace(tableName.toString()));
                                tableLrVo.setBjMoney(org.springframework.util.StringUtils.trimAllWhitespace(bjlj.toString().trim()));
                                tableLrVo.setBnMoney(org.springframework.util.StringUtils.trimAllWhitespace(bnlj.toString().trim()));
                                list.add(tableLrVo);
                            }else if(!newCellValue.equals("行次") && !newCellValue.equals("") && hcIndex.contains(k) && type == 3){
                                TableXjVo tableXjVo = new TableXjVo();
                                Object tableName = getCellValue(row.getCell(k - 1));
                                Object bjlj = getCellValue(row.getCell(k + 1));
                                Object bnlj = getCellValue(row.getCell(k + 2));
                                tableXjVo.setItemName(org.springframework.util.StringUtils.trimAllWhitespace(tableName.toString()));
                                tableXjVo.setBqMoney(org.springframework.util.StringUtils.trimAllWhitespace(bjlj.toString().trim()));
                                tableXjVo.setBnMoney(org.springframework.util.StringUtils.trimAllWhitespace(bnlj.toString().trim()));
                                list.add(tableXjVo);
                            }

                        }

                    }


                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return list;
    }



    public static void replaceExcel(String filePath, List<TableFzVo> tableFzVos, List<TableLrVo> tableLrVos, List<TableXjVo> tableXjVos, MultipartFile multipartFile) throws Exception {

        File file = null;
        if (filePath != null) {
            file = new File(filePath);

            if (!file.exists()) {
                throw new Exception("Excel 文件地址错误不存在");
            }
        }
        // 存放读取的所有结果
        Workbook workBook = null;
        try {

            if (file != null) {
                if (file.getName().endsWith("xlsx")) {
                    workBook = new XSSFWorkbook(file);
                } else if (file.getName().endsWith("xls")) {
                    workBook = new HSSFWorkbook(new FileInputStream(file));
                }
            } else {
                if (multipartFile.getOriginalFilename().endsWith("xlsx")) {
                    workBook = new XSSFWorkbook(multipartFile.getInputStream());
                } else if (multipartFile.getOriginalFilename().endsWith("xls")) {
                    workBook = new HSSFWorkbook(multipartFile.getInputStream());
                }
            }
            for (int i = 0; i < workBook.getNumberOfSheets(); i++ ){
                // 获取第i页数据
                Sheet sheetAt = workBook.getSheetAt(i);

                for(int j =1 ; j < sheetAt.getPhysicalNumberOfRows(); j++){
                    //获取第j行数据
                    Row row = sheetAt.getRow(j);
                    if(Objects.isNull(row)){
                        //如果整行为空， 则跳过
                        continue;
                    }
                    for (int k = 0; k < row.getLastCellNum(); k++) {

                        Object cellValue = getCellValue(row.getCell(k));

                        if(cellValue != null) {
                            String newCellValue = org.springframework.util.StringUtils.trimAllWhitespace(cellValue.toString());
                            if(i == 0){
                                //资产负债表
                                List<TableFzVo> collect = tableFzVos.stream().filter(e -> e.getItemName().trim().equals(newCellValue)).collect(Collectors.toList());
                                if(collect.size() != 0){
                                    TableFzVo tableFzVo = collect.get(0);
                                    if(tableFzVo.getNcs().equals("=ASC(\"ZWB10011012期末余额00\")")){
                                        row.getCell(k + 2).setCellValue("");
                                    }else{
                                        row.getCell(k + 2).setCellValue(Double.valueOf(tableFzVo.getNcs().replaceAll(",","")));
                                    }

                                    if(tableFzVo.getNms().equals("")){
                                        row.getCell(k + 3).setCellValue("");
                                    }else{
                                        row.getCell(k + 3).setCellValue(Double.valueOf(tableFzVo.getNms().replaceAll(",", "")));
                                    }

                                }
                            }else if(i == 1){
                                //利润表
                                List<TableLrVo> collect = tableLrVos.stream().filter(e -> e.getItemName().trim().equals(newCellValue.trim())).collect(Collectors.toList());
                                if(collect.size() != 0){
                                    TableLrVo tableLrVo = collect.get(0);
                                    if(tableLrVo.getBnMoney().equals("")){
                                        row.getCell(k + 2).setCellValue(tableLrVo.getBnMoney());
                                    }else{
                                        row.getCell(k + 2).setCellValue(Double.valueOf(tableLrVo.getBnMoney().replaceAll(",", "")));
                                    }

                                }
                            }else {

                                List<TableXjVo> collect = tableXjVos.stream().filter(e -> e.getItemName().trim().equals(newCellValue.trim())).collect(Collectors.toList());
                                if(collect.size() != 0){
                                    System.out.println(newCellValue);
                                    TableXjVo tableXjVo = collect.get(0);
                                    if(tableXjVo.getBnMoney().equals("")){
                                        row.getCell(k + 2).setCellValue(tableXjVo.getBnMoney());
                                    }else{
                                        row.getCell(k + 2).setCellValue(Double.valueOf(tableXjVo.getBnMoney().replaceAll(",", "")));
                                    }
                                }
                            }
                        }
                    }


                }


            }

            FileOutputStream fileOutputStream = new FileOutputStream("D:/fileOpera/template/mycreate.xlsx");

            workBook.write(fileOutputStream);

            fileOutputStream.flush();
            fileOutputStream.close();



        }catch (Exception e){
            e.printStackTrace();
        }

    }



}
