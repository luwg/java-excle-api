import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;

/**
 * Created by luweigang on 16/8/25.
 */
public class Main {
    public static void main(String[] args) throws IOException, BiffException {
        Workbook workbook = Workbook.getWorkbook(new File("55.xls"));
        Sheet sheet = workbook.getSheet(0);

        Cell a1 = sheet.getCell(1,19);
        Cell b2 = sheet.getCell(2,19);
        Cell c2 = sheet.getCell(3,19);

        String stringa1 = a1.getContents();
        String stringb2 = b2.getContents();
        String stringc2 = c2.getContents();

        System.out.println(stringa1);
        System.out.println(stringb2);
        System.out.println(stringc2);
    }
}
