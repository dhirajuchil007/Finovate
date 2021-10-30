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

public class AddUserTarnsactions {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook("UserTransactions.xlsx");

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
                    if (j == 8 || j == 9 || j == 10 || j == 12 || j == 13)
                        r.getCell(j).setCellType(CellType.STRING);

                }
                OkHttpClient client = new OkHttpClient().newBuilder()
                        .build();
                MediaType mediaType = MediaType.parse("text/plain");
                RequestBody body = RequestBody.create(mediaType, "");
                String Dev_base_URL = "https://sougata.in/cashrich/createCustomerInvestment.json?uname=Us3rname@Dev&pwd=D3vP@ssword2020";
                String SMoke_base_url = "https://sougatabasu.com/cashrich/createCustomerInvestment.json?uname=User@sm0k3&pwd=$mokeP@ss2020";
                String cron_base_url = "https://cron.cashrichapp.in/cashrich/createCustomerInvestment.json?uname=pR0D@U53r!2o2o&pwd=Pr0d@Pa55!2o2o";
                String url = cron_base_url +
                        "&customer_id=" + (int) r.getCell(0).getNumericCellValue() +
                        "&no_unit=" + r.getCell(1).getNumericCellValue() +
                        "&cust_operation=" + (int) r.getCell(2).getNumericCellValue() +
                        "&cust_investment_status=" + (int) r.getCell(3).getNumericCellValue() +
                        "&scheme_master_id=" + (long) r.getCell(4).getNumericCellValue() +
                        "&investment_amount=" + r.getCell(5).getNumericCellValue() +
                        "&investment_plan_type=" + (int) r.getCell(6).getNumericCellValue() +
                        "&allotted_NAV=" + r.getCell(7).getNumericCellValue() +
                        "&folio_no=" + r.getCell(8).getStringCellValue() +
                        "&mf_platform_order_id=" + r.getCell(9).getStringCellValue() +
                        "&remarks=" + r.getCell(10).getStringCellValue() +
                        "&record_type=" + (int) r.getCell(11).getNumericCellValue() +
                        "&create_time=" + r.getCell(12).getStringCellValue() +
                        "&update_time=" + r.getCell(12).getStringCellValue();

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
               // JSONObject responseObj = jsonObject.getJSONObject("response");
              //  String customerInvId = responseObj.getString("custInvestmentId");
                String code = status.getString("code");
                r.createCell(14).setCellValue(code);
                r.createCell(15).setCellValue(output);
             //   r.createCell(16).setCellValue(customerInvId);
               /* if (!code.equals("200"))
                    break;*/
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        wb.write(fout);
        fout.close();
        wb.close();
    }
}
