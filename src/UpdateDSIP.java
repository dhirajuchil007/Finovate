import okhttp3.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class UpdateDSIP {
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook("sip_tobecanceled.xlsx");

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

               /* for (int j = 0; j <= 12; j++) {
                    if (r.getCell(j) == null)
                        r.createCell(j);
                    if (j == 8 || j == 9 || j == 10 || j == 12)
                        r.getCell(j).setCellType(CellType.STRING);

                }*/
                OkHttpClient client = new OkHttpClient().newBuilder()
                        .build();
                MediaType mediaType = MediaType.parse("text/plain");
                RequestBody body = RequestBody.create(mediaType, "");
                String Dev_base_URL = "https://sougata.in/cashrich/cancelSIPOrder.json?uname=Us3rname@Dev&pwd=D3vP@ssword2020";
                String SMoke_base_url = "https://sougatabasu.com/cashrich/cancelSIPOrder.json?uname=User@sm0k3&pwd=$mokeP@ss2020";
                String cron_base_url = "https://cron.cashrichapp.in/cashrich/cancelSIPOrder.json?uname=pR0D@U53r!2o2o&pwd=Pr0d@Pa55!2o2o";


                String url = cron_base_url
                        + "&customer_id=" + (long) r.getCell(2).getNumericCellValue() +
                        "&sip_order_id=" + r.getCell(1).getStringCellValue().substring(1);

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
                JSONObject status = jsonObject.getJSONObject("status");
                String code = status.getString("code");
                r.createCell(7).setCellValue(code);
                r.createCell(8).setCellValue(output);
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
