import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;

/**
 * Created by luweigang on 16/8/25.
 */
public class CreatXL {
    public static String outputFile = "test.xlsx";

    public static void main(String argv[]) {
        try {
            // 创建新的Excel 工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 在Excel工作簿中建一工作表，其名为缺省值
            // 如要新建一名为"效益指标"的工作表，其语句为：
            // HSSFSheet sheet = workbook.createSheet("效益指标");
            HSSFSheet sheet = workbook.createSheet();
            // 在索引0的位置创建行（最顶端的行）
            for (int i=1;i<10;i++){
                for (int j=1;j<10;j++){
                    HSSFRow row = sheet.createRow(j-1);
                    //在索引0的位置创建单元格（左上端）
                    HSSFCell cell = row.createCell(i-1);
                    // 定义单元格为字符串类型
                    cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                    // 在单元格中输入一些内容
                    cell.setCellValue(i*j);
                }
            }
//            HSSFRow row = sheet.createRow(0);
//            //在索引0的位置创建单元格（左上端）
//            HSSFCell cell = row.createCell(0);
//            // 定义单元格为字符串类型
//            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
//            // 在单元格中输入一些内容
//            cell.setCellValue("增加值33");
            // 新建一输出文件流
            FileOutputStream fOut = new FileOutputStream(outputFile);
            // 把相应的Excel 工作簿存盘
            workbook.write(fOut);
            fOut.flush();
            // 操作结束，关闭文件
            fOut.close();
            System.out.println("文件生成...");
        } catch (Exception e) {
            System.out.println("已运行 xlCreate() : " + e);
        }
    }
}
