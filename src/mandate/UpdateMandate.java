package mandate;

import okhttp3.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Scanner;

public class UpdateMandate {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook("update_mandate.xlsx");

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


                OkHttpClient client = new OkHttpClient().newBuilder()
                        .build();
                MediaType mediaType = MediaType.parse("text/plain");
                RequestBody body = RequestBody.create(mediaType, "");
                String Dev_base_URL = "https://sougata.in/cashrich/updateCustomerMandate.json?uname=Us3rname@Dev&pwd=D3vP@ssword2020";
                String SMoke_base_url = "https://sougatabasu.com/cashrich/updateCustomerMandate.json?uname=User@sm0k3&pwd=$mokeP@ss2020";
                String cron_base_url = "https://cron.cashrichapp.in/cashrich/updateCustomerMandate.json?uname=pR0D@U53r!2o2o&pwd=Pr0d@Pa55!2o2o";

                /*String startDateStr = r.getCell(8).getStringCellValue();
                String endDateStr = r.getCell(9).getStringCellValue();

                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy");
                Date startDate = simpleDateFormat.parse(startDateStr);
                Date endDate = simpleDateFormat.parse(endDateStr);
                SimpleDateFormat yyyymmdd = new SimpleDateFormat("dd-MM-yyyy");*/

                String url = cron_base_url
                        + "&customer_mandate_id=" + (long) r.getCell(0).getNumericCellValue()
                        +"&status=13"                        ;
                System.out.println(url);

                Request request = new Request.Builder()
                        .url(url)
                        .method("POST", body)
                        .addHeader("Cookie", "JSESSIONID=YgvlvVPJxQ8P6Th8HWf_NTtau2yXVLtMdtmExH16.localhost")
                        .build();
                Response response = client.newCall(request).execute();
                String output = response.body().string();
                System.out.println(i + " " + output);
                JSONObject jsonObject = new JSONObject(output);
                JSONObject responseObj=jsonObject.getJSONObject("response");
                JSONObject status = jsonObject.getJSONObject("status");
                String code = status.getString("code");
                r.createCell(3).setCellValue(code);
                r.createCell(4).setCellValue(output);
                r.createCell(5).setCellValue(responseObj.getString("customerMandateId"));
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
