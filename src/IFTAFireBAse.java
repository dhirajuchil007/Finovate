import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;

import com.fasterxml.jackson.databind.ObjectMapper;

import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;

public class IFTAFireBAse {

	public static void main(String[] args) throws IOException {
		
		  // TODO Auto-generated method stub 
		Workbook wb=new XSSFWorkbook("users.xlsx");
		  
		  Sheet s=wb.getSheetAt(0);
		  
		  
		  FileOutputStream fout=new FileOutputStream("out.xlsx");

		
		try {
		int n=6;
		Row r=s.getRow(1);
		Cell c =r.getCell(0);
	
		
		for(int i=53;i<=57;i++) {
			//System.out.println(i);
			r=s.getRow(i);
			c=r.getCell(10);
			double x=c.getNumericCellValue();
			if(x==1) {
				
				c=r.getCell(8);
				c.setCellType(CellType.STRING);
				String cust_id=c.getStringCellValue();
				
					c=r.getCell(5);					
						
	String url=c.getStringCellValue();
	
	c=r.getCell(0);
	String fname=c.getStringCellValue();
	
	c=r.getCell(1);
	String lname=c.getStringCellValue();
	
	c=r.getCell(2);
	String email=c.getStringCellValue();
	
	c=r.getCell(3);
	String designation=c.getStringCellValue();
	
	
	c=r.getCell(4);
	String company=c.getStringCellValue();
	
	ObjectMapper mapper = new ObjectMapper();
	
	User user=new User();
	user.setFirst_name(fname);
	user.setLast_name(lname);
	user.setEmail(email);
	user.setImage_url(url);
	user.setCompany(company);
	user.setDesignation(designation);
	user.setCust_id(cust_id);
	
				OkHttpClient client = new OkHttpClient().newBuilder()
						  .build();
						MediaType mediaType = MediaType.parse("application/x-www-form-urlencoded");
						
						String json= mapper.writerWithDefaultPrettyPrinter().writeValueAsString(user);
						
						System.out.print(json);
						RequestBody body = 
								RequestBody.create(mediaType, json);
						Request request = new Request.Builder()
						  .url("https://zoho-subscription-ifta-2020.firebaseio.com/registerdUsersList.json?auth=BrVw3cpJLdJx74W3MDb68xlvQBgEealxVQWYmrcd")
						  .method("POST", body)
						  .addHeader("Accept", "application/json")
						  .addHeader("Content-Type", "application/x-www-form-urlencoded")
							/*
							 * .addHeader("appId", "263517b7071bf4c") .addHeader("apiKey",
							 * "18902cb47931426ed749d797e910ededa0394b34")
							 */
						  .build();
						Response response = client.newCall(request).execute();
						
						String message=response.body().string();
						System.out.println(i+" "+message);
						c=r.createCell(12);
						c.setCellValue(message);
			
			}
			
		
		}
		}catch(Exception e) {
			e.printStackTrace();
			wb.write(fout);
			fout.close();
			wb.close();
		}
		
		
		wb.write(fout);
		fout.close();
		wb.close();

	}

}
