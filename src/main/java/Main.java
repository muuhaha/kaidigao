import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;

/**
 * Created by me on 2017/2/26.
 */
public class Main {

    static String[] titles = {"月","日","字","号","摘要","借方金额","贷方金额","对方科目"};
    static String[] outputTitles = {"日期","记账凭证","摘要","借方金额","贷方金额","对方科目"};

    public static void main(String[] args) throws Exception {
        Scanner in = new Scanner(System.in);

        System.out.println("请输入需要处理的文件路径（例如：F:\\\\xg\\\\kaidigao\\\\doc\\\\示例.xls）：");
//        String path = in.nextLine();
        String path = "F:\\xg\\kaidigao\\doc\\示例.xls";

        System.out.println("请输入需要处理的 sheet 的编号，从1开始，以空格间隔（例如：1 2 3）：");
        String sheetStr = in.nextLine();
        sheetStr = sheetStr.trim();
        String[] strs = sheetStr.split(" ");
        int[] sheetIndexes = new int[strs.length];
        for (int i = 0; i < sheetIndexes.length; i++) {
            sheetIndexes[0] = Integer.valueOf(strs[i]) - 1;
        }

//        backUpOriginFile(path);
        func(path, sheetIndexes);
    }

    public static void backUpOriginFile(String path) {
        Utils.copyFile(path, "doc/备份-示例1.xlsx");
    }

    public static void initGUI() {
        //实例化一个JFrame类的对象
        javax.swing.JFrame jf = new javax.swing.JFrame();
        //设置窗体的标题属性
        jf.setTitle("开底稿");
        //设置窗体的大小属性
        jf.setSize(250,350);
        //设置窗体的位置属性
        jf.setLocation(400,200);
        //设置窗体关闭时退出程序
        jf.setDefaultCloseOperation(3);
        //设置禁止调整窗体的大小
        jf.setResizable(false);
        //设置窗体是否可见
        jf.setVisible(true);
    }

    public static void func(String path, int[] sheetIndexes) throws Exception {
        Workbook book = new XSSFWorkbook(path);

        for (int sheetIndex : sheetIndexes) {

            Sheet defaultSheet = book.getSheetAt(sheetIndex);

            System.out.println("处理编号为：" + sheetIndex + "的sheet, sheet名为：" + defaultSheet.getSheetName());

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

            Scanner in = new Scanner(System.in);
            System.out.println("请输入金额阈值，小于该值将被过滤（例如：1000）：");
            double minMoney = in.nextDouble();

            // 遍历sheet，查找符合条件的row
            List<Map<String, Cell>> valids = new ArrayList<>();
            for (int row = maxRowIndex + 1; defaultSheet.getRow(row) != null; row++) {
                if (defaultSheet.getRow(row).getCell(titleCells.get("借方金额").getColumnIndex()).getNumericCellValue()
                        >= minMoney ||
                        defaultSheet.getRow(row).getCell(titleCells.get("贷方金额").getColumnIndex()).getNumericCellValue()
                                >= minMoney) {
                    if (defaultSheet.getRow(row).getCell(titleCells.get("摘要").getColumnIndex()).getStringCellValue().equals("结转损益")) {
                        continue;
                    }
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


            // 截取部分条目: number
            System.out.println("保留金额最大的n条记录，请输入n值，（例如：20）：");
            int number = in.nextInt();

            double thred = 0;
            if (valids.size() > number) {
                List<Double> moneys = new ArrayList<>();
                for (Map<String, Cell> tmp : valids) {
                    moneys.add(tmp.get("借方金额").getNumericCellValue());
                    moneys.add(tmp.get("贷方金额").getNumericCellValue());
                }
                Collections.sort(moneys, new Comparator<Double>() {
                    @Override
                    public int compare(Double o1, Double o2) {
                        if (o1 > o2) return -1;
                        if (o1 < o2) return 1;
                        return 0;
                    }
                });
                thred = moneys.get(number - 1);
                List<Map<String, Cell>> filtered = new ArrayList<>();
                for (Map<String, Cell> tmp : valids) {
                    if (tmp.get("借方金额").getNumericCellValue() >= thred || tmp.get("贷方金额").getNumericCellValue() >= thred) {
                        filtered.add(tmp);
                    }
                }
                valids = filtered;
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
        OutputStream out = new FileOutputStream(path + "A");
        book.write(out);
        System.out.println("完成！");
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
