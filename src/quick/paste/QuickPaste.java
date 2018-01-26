/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quick.paste;

import java.awt.AWTException;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Vector;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Bluemond
 */
public class QuickPaste {

    /**
     *
     * @param args
     *
     * @throws AWTException
     *
     */
    public Vector<String> code = new Vector<String>();

    public static void main(String[] args) throws Exception {
        Form f = new Form();
        // TODO code application logic here

    }

    public void testReadExcel(String addr) {
        try {
            // 读取Excel  
            InputStream is = new FileInputStream(addr);
            Workbook wb = new XSSFWorkbook(is);

            // 获取sheet数目  
            Sheet sheet = wb.getSheetAt(0);
            Row row = null;
            int lastRowNum = sheet.getLastRowNum();
            // 循环读取  
            for (int i = 0; i <= lastRowNum; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    // 获取每一列的值

                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        Cell cell = row.getCell(j);
                        String value = getCellValue(cell);
                        if (!value.equals("")) {
                            code.add(value);

                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private String getCellValue(Cell cell) {
        Object result = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    result = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    long num = (long) cell.getNumericCellValue();
                    result = num;
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    result = cell.getBooleanCellValue();
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    result = cell.getCellFormula();
                    break;
                case Cell.CELL_TYPE_ERROR:
                    result = cell.getErrorCellValue();
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
                default:
                    break;
            }
        }
        return result.toString();
    }


}
