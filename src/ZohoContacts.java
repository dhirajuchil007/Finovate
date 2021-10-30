import me.xdrop.fuzzywuzzy.FuzzySearch;
import okhttp3.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

public class ZohoContacts {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook("zohoConvert.xlsx");

        Sheet s = wb.getSheetAt(0);


        FileOutputStream fout = new FileOutputStream("zohoout.xlsx");

        Scanner sc = new Scanner(System.in);
        System.out.println("Enter start index");
        int start = sc.nextInt();
        System.out.println("enter end index");
        int end = sc.nextInt();
        Row r = s.getRow(1);
        try {


            for (int i = start; i <= end; i++) {
                Thread.sleep(500);
                r = s.getRow(i);
                r.getCell(0).setCellType(CellType.STRING);
                String customerId = r.getCell(0).getStringCellValue();
                OkHttpClient client = new OkHttpClient().newBuilder().connectTimeout(20, TimeUnit.SECONDS)
                        .build();
                MediaType mediaType = MediaType.parse("text/plain");
                RequestBody body = RequestBody.create(mediaType, "");
                Request request = new Request.Builder()
                        .url("https://cron.cashrichapp.in/cashrich//convertZohoLeadToContact.json?uname=pR0D@U53r!2o2o&pwd=Pr0d@Pa55!2o2o&customerId=" + customerId)
                        .method("POST", body)
                        .build();
                Response response = client.newCall(request).execute();
                String output = response.body().string();
                JSONObject jsonObject = new JSONObject(output);
                JSONObject status = jsonObject.getJSONObject("status");
                String code = status.getString("code");

                r.createCell(1).setCellValue(code);
                r.createCell(2).setCellValue(output);
                System.out.println(i + " " + output);

            }
        } catch (Exception e) {
            e.printStackTrace();
        }


        wb.write(fout);
        fout.close();
        wb.close();
    }
}
