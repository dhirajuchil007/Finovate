
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.util.JSONPObject;
import okhttp3.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;
import userAdd.AddUserRequest;
import userAdd.AddressDetails;
import userAdd.BankDetails;
import userAdd.FatcaDetails;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Scanner;


public class BulkAddUsers {
    // TODO Auto-generated method stub
    public static void main(String[] args) throws IOException {
        Workbook wb = new XSSFWorkbook("addUsers.xlsx");

        Sheet s = wb.getSheetAt(0);


        FileOutputStream fout = new FileOutputStream("out.xlsx");

        Scanner sc = new Scanner(System.in);
        System.out.println("Enter start index");
        int start = sc.nextInt();
        System.out.println("enter end index");
        int end = sc.nextInt();
        Row r = s.getRow(1);
        try {
            for (int i = start; i < end; i++) {
                r = s.getRow(i);
                for (int j = 0; j <= 30; j++) {
                    Cell c = r.getCell(j);
                    if (c != null) {
                        if (j == 4 || j == 5 || j == 14 || j == 22 || j == 23 || j == 25 | j == 26 || j == 29 || j == 30)
                            continue;
                        //  System.out.print(j);
                        c.setCellType(CellType.STRING);
                    } else {
                        r.createCell(j);
                    }
                }

                int fuzzyMatch = (int) r.getCell(30).getNumericCellValue();
                if (fuzzyMatch < 50) {
                    r.createCell(31).setCellValue("Skipped:Fuzzy match below threshold");
                    System.out.println("Skipped:Fuzzy match below threshold");
                    continue;
                }
                AddUserRequest adduserRequest = new AddUserRequest();
                adduserRequest.setFirstName(r.getCell(0).getStringCellValue())
                        .setLastName(r.getCell(1).getStringCellValue())
                        .setCountryCode(r.getCell(2).getStringCellValue())
                        .setMobileNo(r.getCell(3).getStringCellValue())
                        .setGender((int) r.getCell(4).getNumericCellValue())
                        .setDateOfBirth(r.getCell(5).getDateCellValue())
                        .setEmailId(r.getCell(6).getStringCellValue())
                        .setPan(r.getCell(7).getStringCellValue())
                        .setReferalNumber("54321")
                        .setNameOnPan(r.getCell(8).getStringCellValue());
//                    .setFatherSpouseName(r.getCell(9).getStringCellValue())
//                    .setMotherName(r.getCell(10).getStringCellValue());

                BankDetails bankDetails = new BankDetails();
                bankDetails.setAccountType(r.getCell(11).getStringCellValue())
                        .setAccountNo(r.getCell(12).getStringCellValue())
                        .setIfscCode(r.getCell(13).getStringCellValue())
                        .setBankCode((int) r.getCell(14).getNumericCellValue())
                        .setNameInBank(r.getCell(15).getStringCellValue())
                        .setIsDefault(true).setValidationCode(10004);

                adduserRequest.setBankDetails(bankDetails);

                AddressDetails addressDetails = new AddressDetails();
                addressDetails.setAddressLine(r.getCell(16) + " " + r.getCell(17))
                        .setCity(r.getCell(18).getStringCellValue())
                        .setState(r.getCell(19).getStringCellValue())
                        .setCountry(r.getCell(20).getStringCellValue())
                        .setPinCode(r.getCell(21).getStringCellValue());

                adduserRequest.setAddressDetails(addressDetails);

                FatcaDetails fatcaDetails = new FatcaDetails();
                fatcaDetails.setTaxStatus((int) r.getCell(22).getNumericCellValue())
                        .setAnnualIncome((int) r.getCell(23).getNumericCellValue())
                        .setNationality((int) r.getCell(25).getNumericCellValue())
                        .setTaxResidentCountry((int) r.getCell(26).getNumericCellValue())
                        .setPoliticallyExposed(false)
                        .setFatcaCompliant(true);

                adduserRequest.setFatcaDetails(fatcaDetails);

                adduserRequest.setNomineeName(r.getCell(28).getStringCellValue());
                adduserRequest.setNomineeRelation((int) r.getCell(29).getNumericCellValue());

                //activate account-----------------------------------------------------
                adduserRequest.setToBeActivated(false);

                ObjectMapper objectMapper = new ObjectMapper();
                final DateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
                objectMapper.setDateFormat(df);
                String jsonString = objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(adduserRequest);

                //  System.out.println(jsonString);

                OkHttpClient client = new OkHttpClient().newBuilder()
                        .build();
                MediaType mediaType = MediaType.parse("application/json");
                RequestBody body = RequestBody.create(mediaType, jsonString);
                String smokeUrl = "https://sougatabasu.com/cashrich//addNewCustomer.json?uname=user@sm0Ke&pwd=sm0k3P@ss2022";
                String cronUrl = "https://cashrichapp.in/cashrich//addNewCustomer.json?uname=u53R@PR0d!2!2!&pwd=Pa55@PR0d!2!2!";
                Request request = new Request.Builder()
                        .url(smokeUrl)
                        .method("POST", body)
                        .addHeader("Content-Type", "application/json")
                        .build();
                Response response = client.newCall(request).execute();
                String output = response.body().string();
                System.out.println(i + " " + output);
                JSONObject jsonObject = new JSONObject(output);
                JSONObject status = jsonObject.getJSONObject("status");
                String code = status.getString("code");
                r.createCell(31).setCellValue(output);
                r.createCell(32).setCellValue(code);

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        wb.write(fout);
        fout.close();
        wb.close();
    }


    public interface ColumnsList {

        int FNAME = 0;

    }

}
