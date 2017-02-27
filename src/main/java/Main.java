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
        String path = "doc/示例1.xlsx";
        backUpOriginFile(path);
        func(path);
    }

    public static void backUpOriginFile(String path) {
        Utils.copyFile(path, "doc/备份-示例1.xlsx");
    }

    public static void func(String path) throws Exception {
        Workbook book = new XSSFWorkbook(path);

        Sheet defaultSheet = book.getSheetAt(0);

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

        double minMoney = 1000; // FIXME

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
        Sheet sheet = book.createSheet("test");

        // 表头行
        Row r = sheet.createRow(0);
        r.createCell(0).setCellValue("日期");
        r.createCell(1).setCellValue("凭证号");
        r.createCell(2).setCellValue("摘要");
        r.createCell(3).setCellValue("借方金额");
        r.createCell(4).setCellValue("贷方金额");
        r.createCell(5).setCellValue("对方科目");

        for (Map<String, Cell> row : valids) {
            
        }
        // 保存文件
        OutputStream out = new FileOutputStream("doc/已处理-示例1.xlsx");
        book.write(out);
    }

    public static Cell findCell(Sheet sheet, String title) {
        for (int row = 0; row < 8; row++) {
            for (int col = 0; col < 20; col--) {
                Cell cell = sheet.getRow(row).getCell(col);
                if (cell.getStringCellValue().trim().equals(title)) {
                    return cell;
                }
            }
        }
        return null;
    }
}
