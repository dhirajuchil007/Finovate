import me.xdrop.fuzzywuzzy.FuzzySearch;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class FuzzyCheck {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook("Fuzzy.xlsx");

        Sheet s = wb.getSheetAt(0);


        FileOutputStream fout = new FileOutputStream("out.xlsx");

        Scanner sc = new Scanner(System.in);
        System.out.println("Enter start index");
        int start = sc.nextInt();
        System.out.println("enter end index");
        int end = sc.nextInt();
        Row r = s.getRow(1);
        try {


            for (int i = start; i <= end; i++) {
                r = s.getRow(i);
                String nameOnPan = "";
                String nameInBank = "";
                if (r.getCell(0) != null)
                    nameOnPan = r.getCell(0).getStringCellValue();
                if (r.getCell(1) != null)
                    nameInBank = r.getCell(1).getStringCellValue();
                System.out.println(i);
                int percentage = FuzzySearch.tokenSortRatio(nameOnPan, nameInBank);
                nameOnPan = nameOnPan.replaceAll(" ", "");
                nameInBank = nameInBank.replaceAll(" ", "");
                percentage = Math.max(percentage, FuzzySearch.tokenSortRatio(nameOnPan, nameInBank));
                r.createCell(3).setCellValue(percentage);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }


        wb.write(fout);
        fout.close();
        wb.close();
    }
}
