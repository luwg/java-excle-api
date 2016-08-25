import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;

/**
 * Created by luweigang on 16/8/25.
 */
public class ReadXL {
    public static String fileToBeRead="test.xlsx";
    public static void main(String argv[]){
        try{
            // 创建对Excel工作簿文件的引用
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileToBeRead));
            // 创建对工作表的引用。
            // 本例是按名引用（让我们假定那张表有着缺省名"Sheet0"）
            HSSFSheet sheet = workbook.getSheet("Sheet0");
            // 也可用getSheetAt(int index)按索引引用，
            // 在Excel文档中，第一张工作表的缺省索引是0，
            // 其语句为：HSSFSheet sheet = workbook.getSheetAt(0);
            // 读取左上端单元
            HSSFRow row = sheet.getRow(0);
            HSSFCell cell = row.getCell(0);
            // 输出单元内容，cell.getStringCellValue()就是取所在单元的值
            System.out.println("左上端单元是： " + cell.getStringCellValue());
        }catch(Exception e) {
            System.out.println("已运行xlRead() : " + e );
        }
    }
}
