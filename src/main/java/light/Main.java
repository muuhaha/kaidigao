package light;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.List;

/**
 * Created by me on 2017/2/26.
 */
public class Main  extends JFrame implements ActionListener{

    static String[] titles = {"月","日","字","号","摘要","借方金额","贷方金额","对方科目"};
    static String[] outputTitles = {"日期","记账凭证","摘要","借方金额","贷方金额","对方科目"};
    static int[] sheetWidths = {14, 8, 30, 12, 12, 20};

    public static void main(String[] args) throws Exception {
        new Main().initGUI();
    }

    public static void backUpOriginFile(String path) {
        Utils.copyFile(path, "doc/备份-示例1.xlsx");
    }

    public static void initProcess(File file) throws Exception {
        Scanner in = new Scanner(System.in);

        // 命令行方式读取文件路径=========
        if (!file.exists() || !file.isFile()) {
            System.out.println("请手动输入需要处理的文件路径（例如：F:\\xg\\kaidigao\\doc\\示例.xls）：");
            String path = in.nextLine();
            file = new File("D:\\work\\kaidigao\\doc\\示例1.xlsx");
        }
        // ===============================

        System.out.println("请输入需要处理的 sheet 的编号，从1开始，以空格间隔（例如：1 2 3）：");
        String sheetStr = in.nextLine();
        sheetStr = sheetStr.trim();
        String[] strs = sheetStr.split(" ");
        int[] sheetIndexes = new int[strs.length];
        for (int i = 0; i < sheetIndexes.length; i++) {
            sheetIndexes[i] = Integer.valueOf(strs[i]) - 1;
        }

        func(file, sheetIndexes);
    }

    public static void func(File file, int[] sheetIndexes) throws Exception {
        String fileName = file.getName();
        String suffix = fileName.substring(fileName.length() - 4);
        Workbook book = null;
        if (suffix.equals(".xls")) {
            book = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(file)));
        } else if (suffix.equals("xlsx")) {
            book = new XSSFWorkbook(file);
        } else {
            throw new Exception();
        }

        for (int sheetIndex : sheetIndexes) {

            Sheet defaultSheet = book.getSheetAt(sheetIndex);

            System.out.println("处理编号为：" + (sheetIndex + 1) + " 的sheet, sheet名为：" + defaultSheet.getSheetName());

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
                sheet.setColumnWidth(i, sheetWidths[i] * 256);
                outputColIndex.put(outputTitles[i], i);
            }


            // 截取部分条目: number
            System.out.println("保留金额最大的n条记录，0或者不输入表示全都保留，（例如：20）：");
            in.nextLine();
            String numberStr = in.nextLine();
            int number;
            if (numberStr == null || numberStr.trim().length() == 0) {
                number = 0;
            } else {
                number = Integer.valueOf(numberStr.trim());
            }

            double thred = 0;
            if (number != 0 && valids.size() > number) {
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
        OutputStream out = new FileOutputStream(file.getParent() + "\\" + "已处理-" + file.getName());
        book.write(out);
        out.close();
        System.out.print("完成！");
        System.out.println("保存为：" + file.getParent() + "\\" + "已处理-" + file.getName());
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

    JFileChooser fc = new JFileChooser();//创建文件对话框对象
    JButton open;

    public void initGUI()
    {
        Container container = this.getContentPane();
        container.setLayout(new FlowLayout());
        this.setTitle("打开底稿文件");
        open = new JButton("打开文件");
        open.addActionListener(this);
        container.add(open);//添加到内容窗格上
        this.setVisible(true);
        this.setSize(200,80);

    }

    public void actionPerformed(ActionEvent e)
    {
        JButton button = (JButton)e.getSource();//得到事件源
        if(button == open)//选择的是“打开文件”按钮
        {
            int select = fc.showOpenDialog(this);//显示打开文件对话框
            if(select == JFileChooser.APPROVE_OPTION)//选择的是否为“确认”
            {
                File file = fc.getSelectedFile();
                System.out.println("选择文件："+file.getName());
                try {
                    initProcess(file);
                } catch (Exception ex) {
                    System.out.println("处理出现错误！");
                    ex.printStackTrace();
                }
                System.out.println("请选择下一个需要处理的文件！");
            }
            else
                System.out.println("操作被取消");
        }
    }
}
