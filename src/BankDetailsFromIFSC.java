import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class BankDetailsFromIFSC {

    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook("IFSC.xlsx");

        Sheet s = wb.getSheetAt(0);


        FileOutputStream fout = new FileOutputStream("trans_out.xlsx");
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter start index");
        int start = sc.nextInt();
        System.out.println("enter end index");
        int end = sc.nextInt();
        Row r = s.getRow(1);

        try {
            for (int i = start; i <= end; i++) {
                r = s.getRow(i);
                String ifsc = r.getCell(0).getStringCellValue();

                OkHttpClient client = new OkHttpClient().newBuilder()
                        .build();
                Request request = new Request.Builder()
                        .url("https://ifsc.razorpay.com/"+ifsc)
                        .method("GET", null)
                        .build();
                Response response = client.newCall(request).execute();

                String output = response.body().string();
                System.out.println(i+" "+output);
                JSONObject jsonObject = new JSONObject(output);
                String branch=jsonObject.getString("BRANCH");
                String BANK=jsonObject.getString("BANK");

                r.createCell(2).setCellValue(branch);
                r.createCell(1).setCellValue(BANK);

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        wb.write(fout);
        fout.close();
        wb.close();


    }
}
