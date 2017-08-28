package com.gilda.tools;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSplit {
    public static void main(String[] args) {
        try {
            System.out.println("开始拆分.....");
            Map<String, XSSFWorkbook> map = getSplitMap("D:/xxxx.xlsx");// 得到拆分后的子文件存储对象
            createSplitXSSFWorkbook(map, "D:/splitdone/201404", "2014.04");// 遍历对象生成的拆分文件
            System.out.println("拆分结束,文件被拆分为" + map.size() + "个文件.");
        } catch (Exception e) {

            e.printStackTrace();
        }
    }

    // 将第一列的值作为键值,将一个文件拆分为多个文件
    public static Map<String, XSSFWorkbook> getSplitMap(String fileName)
            throws Exception {

        Map<String, XSSFWorkbook> map = new HashMap<String, XSSFWorkbook>();

        InputStream is = new FileInputStream(new File(fileName));
        // 根据输入流创建Workbook对象
        Workbook wb = WorkbookFactory.create(is);
        // get到Sheet对象
        Sheet sheet = wb.getSheetAt(0);
        Row titleRow = null;
        // 这个必须用接口
        int i = 0;
        for (Row row : sheet) {// 遍历每一行
            if (i == 0) {
                titleRow = row;// 得到标题行
            } else {
                Cell keyCell = row.getCell(0);
                String key = keyCell.getRichStringCellValue().toString();
                XSSFWorkbook tempWorkbook = map.get(key);
                if (tempWorkbook == null) {// 如果以当前行第一列值作为键值取不到工作表
                    tempWorkbook = new XSSFWorkbook();
                    Sheet tempSheet = tempWorkbook.createSheet();
                    Row firstRow = tempSheet.createRow(0);
                    for (short k = 0; k < titleRow.getLastCellNum(); k++) {// 为每个子文件创建标题
                        Cell c = titleRow.getCell(k);
                        Cell newcell = firstRow.createCell(k);
                        newcell.setCellValue(c.getStringCellValue());
                    }
                    map.put(key, tempWorkbook);
                }
                Sheet secSheet = tempWorkbook.getSheetAt(0);
                Row secRow = secSheet.createRow(secSheet.getLastRowNum() + 1);
                for (short m = 0; m < row.getLastCellNum(); m++) {
                    Cell newcell = secRow.createCell(m);
                    setCellValue(newcell, row.getCell(m), tempWorkbook);
                }
                map.put(key, tempWorkbook);
            }
            i = i + 1;// 行数加一

        }
        return map;
    }

    // 创建文件
    public static void createSplitXSSFWorkbook(Map<String, XSSFWorkbook> map,
            String savePath, String month) throws IOException {
        Iterator iter = map.entrySet().iterator();
        while (iter.hasNext()) {
            Map.Entry entry = (Map.Entry) iter.next();
            String key = (String) entry.getKey();
            XSSFWorkbook val = (XSSFWorkbook) entry.getValue();
            File filePath = new File(savePath);
            if (!filePath.exists()) {
                System.out.println("存放目录不存在,自动为您创建存放目录.");
                filePath.mkdir();
            }
            if (!filePath.isDirectory()) {
                System.out.println("无效文件目录");
                return;
            }
            File file = new File(savePath + "/" + key + "_" + month + ".xlsx");
            FileOutputStream fOut;// 新建输出文件流
            try {
                fOut = new FileOutputStream(file);
                val.write(fOut); // 把相应的Excel工作薄存盘
                fOut.flush();
                fOut.close(); // 操作结束，关闭文件
            } catch (FileNotFoundException e) {
                System.out.println("找不到文件");
            }
        }
    }

    // 将一个单元格的值赋给另一个单元格
    public static void setCellValue(Cell newCell, Cell cell, XSSFWorkbook wb) {
        if (cell == null) {
            return;
        }
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_BOOLEAN:
            newCell.setCellValue(cell.getBooleanCellValue());
            break;
        case Cell.CELL_TYPE_NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                XSSFCellStyle cellStyle = wb.createCellStyle();
                XSSFDataFormat format = wb.createDataFormat();
                cellStyle.setDataFormat(format.getFormat("yyyy/m/d"));
                newCell.setCellStyle(cellStyle);
                newCell.setCellValue(cell.getDateCellValue());
            } else {
                // 读取数字
                newCell.setCellValue(cell.getNumericCellValue());
            }
            break;
        case Cell.CELL_TYPE_FORMULA:
            newCell.setCellValue(cell.getCellFormula());
            break;
        case Cell.CELL_TYPE_STRING:
            newCell.setCellValue(cell.getStringCellValue());
            break;
        }

    }

}