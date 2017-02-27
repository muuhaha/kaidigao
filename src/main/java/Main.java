import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by me on 2017/2/26.
 */
public class Main {
    static String[] titles = {"月","日","字","号","摘要","借方金额","贷方金额","对方科目"};
    static String[] outputTitles = {"日期","记账凭证","摘要","借方金额","贷方金额","对方科目"};
    public static void main(String[] args) throws Exception {
        String path = "doc/原表.xlsx";
//        backUpOriginFile(path);
        func(path);
    }

    public static void backUpOriginFile(String path) {
        Utils.copyFile(path, "doc/备份-示例1.xlsx");
    }

    public static void func(String path) throws Exception {
        Workbook book = new XSSFWorkbook(path);

        for (int sheetIndex = 0; sheetIndex < 5; sheetIndex++) {

            Sheet defaultSheet = book.getSheetAt(sheetIndex);

            // 查找表头所在 Cell
            HashMap<String, Cell> titleCells = new HashMap<>();
            int maxRowIndex = -1; // 不同表头不在同一行，取最下面的一行
            for (String title : titles) {
                Cell cell = findCell(defaultSheet, title);
                if (cell != null) {
                    titleCells.put(title, cell);
                    maxRowIndex = Math.max(maxRowIndex, cell.getRowIndex());
                } else {
                    throw new Exception("未找到表头：" + title);
                }
            }

            double minMoney = 10000; // FIXME

            // 遍历sheet，查找符合条件的row
            List<Map<String, Cell>> valids = new ArrayList<>();
            for (int row = maxRowIndex + 1; defaultSheet.getRow(row) != null; row++) {
                if (defaultSheet.getRow(row).getCell(titleCells.get("借方金额").getColumnIndex()).getNumericCellValue()
                        >= minMoney ||
                        defaultSheet.getRow(row).getCell(titleCells.get("贷方金额").getColumnIndex()).getNumericCellValue()
                                >= minMoney) {
                    Map<String, Cell> tmp = new HashMap<>();
                    for (String title : titles) {
                        tmp.put(title, defaultSheet.getRow(row).getCell(titleCells.get(title).getColumnIndex()));
                    }
                    valids.add(tmp);
                }
            }

            // 新建结果sheet
            Sheet sheet = book.createSheet(defaultSheet.getSheetName() + "-处理后");

            // 表头行
            Row headRow = sheet.createRow(0);
            Map<String, Integer> outputColIndex = new HashMap<>();
            for (int i = 0; i < outputTitles.length; i++) {
                headRow.createCell(i).setCellValue(outputTitles[i]);
                outputColIndex.put(outputTitles[i], i);
            }

            // 日期格式
            CellStyle cellStyle = book.createCellStyle();
            CreationHelper createHelper = book.getCreationHelper();
            cellStyle.setDataFormat(
                    createHelper.createDataFormat().getFormat("yyyy/m/d"));

            for (int i = 0; i < valids.size(); i++) {
                Map<String, Cell> tmp = valids.get(i);
                Row row = sheet.createRow(i + 1);

                // 日期
                Cell newCell = row.createCell(0);
                newCell.setCellValue(Utils.parseDate("2016-"+tmp.get("月").getStringCellValue()+"-"+tmp.get("日").getStringCellValue()));
                newCell.setCellStyle(cellStyle);

                // 凭证
                newCell = row.createCell(1);
                newCell.setCellValue(tmp.get("字").getStringCellValue()+tmp.get("号").getStringCellValue());

                // 摘要
                newCell = row.createCell(2);
                newCell.setCellValue(tmp.get("摘要").getStringCellValue());

                // 借方金额
                newCell = row.createCell(3);
                newCell.setCellValue(tmp.get("借方金额").getNumericCellValue());

                // 贷方金额
                newCell = row.createCell(4);
                newCell.setCellValue(tmp.get("贷方金额").getNumericCellValue());

                // 对方科目
                newCell = row.createCell(5);
                newCell.setCellValue(tmp.get("对方科目").getStringCellValue());

            }
        }

        // 保存文件
        OutputStream out = new FileOutputStream("doc/原表-已处理.xlsx");
        book.write(out);
    }

    public static Cell findCell(Sheet sheet, String title) {
        for (int row = 0; row < 8; row++) {
            Row r = sheet.getRow(row);
            if (r == null) continue;
            for (int col = 0; col < 20; col++) {
                Cell cell = r.getCell(col);
                if (cell == null) continue;
                if (cell.getStringCellValue().trim().equals(title)) {
                    return cell;
                }
            }
        }
        return null;
    }
}
