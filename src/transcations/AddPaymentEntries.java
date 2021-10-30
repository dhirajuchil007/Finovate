package transcations;

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

public class AddPaymentEntries {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook("UserPayments.xlsx");

        Sheet s = wb.getSheetAt(0);


        FileOutputStream fout = new FileOutputStream("trans_out.xlsx");
        Scanner sc = new Scanner(System.in);
        System.out.println("Enter start index");
        int start = sc.nextInt();
        System.out.println("enter end index");
        int end = sc.nextInt();
        Row r = s.getRow(1);

        try {
            for (int i = end; i >= start; i--) {
                r = s.getRow(i);

                for (int j = 0; j <= 12; j++) {
                    if (r.getCell(j) == null)
                        r.createCell(j);
                    if (j == 0)
                        r.getCell(j).setCellType(CellType.STRING);

                }
                OkHttpClient client = new OkHttpClient().newBuilder()
                        .build();
                MediaType mediaType = MediaType.parse("text/plain");
                RequestBody body = RequestBody.create(mediaType, "");
                String Dev_base_URL = "https://sougata.in/cashrich/createPaymentEntry.json?uname=Us3rname@Dev&pwd=D3vP@ssword2020";
                String SMoke_base_url = "https://sougatabasu.com/cashrich/createPaymentEntry.json?uname=User@sm0k3&pwd=$mokeP@ss2020";
                String cron_base_url = "https://cron.cashrichapp.in/cashrich/createPaymentEntry.json?uname=pR0D@U53r!2o2o&pwd=Pr0d@Pa55!2o2o";
                String url = SMoke_base_url +
                        "&customer_investment_id=" + r.getCell(0).getStringCellValue() +
                        "&amount=" + r.getCell(1).getNumericCellValue() +
                        "&payment_provider=" + (int) r.getCell(2).getNumericCellValue() +
                        "&payment_instrument=" + (int) r.getCell(3).getNumericCellValue() +
                        "&instrument_details=" + "NA" +
                        "&status=" + 142 +
                        "&customer_bank_id=" + r.getCell(5).getStringCellValue() +
                        "&mandate_detail_id=" + r.getCell(6).getStringCellValue();

                System.out.println(url);
                Request request = new Request.Builder()
                        .url(url)
                        .method("POST", body)
                        .build();
                Response response = client.newCall(request).execute();
                String output = response.body().string();
                System.out.println(i + " " + output);
                JSONObject jsonObject = new JSONObject(output);
                JSONObject status = jsonObject.getJSONObject("status");
                String code = status.getString("code");
                r.createCell(14).setCellValue(code);
                r.createCell(15).setCellValue(output);
                if (!code.equals("200"))
                    break;
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        wb.write(fout);
        fout.close();
        wb.close();
    }
}
