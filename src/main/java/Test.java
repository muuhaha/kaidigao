import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created by zhaoxiaoguang on 2017/2/27.
 */
public class Test {
    public static void main(String[] args) throws Exception {
        String path = "doc/示例1.xlsx";
        Workbook book = new XSSFWorkbook(path);
        Sheet defaultSheet = book.getSheetAt(0);
        Row r = defaultSheet.getRow(4000);
        Cell cell = defaultSheet.getRow(4000).getCell(0);
        String d = cell.getStringCellValue();
        System.out.print(d);
    }
}
