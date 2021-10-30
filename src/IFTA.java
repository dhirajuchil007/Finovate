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

import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;

public class IFTA {

	public static void main(String[] args) throws IOException {
		
		  // TODO Auto-generated method stub 
		Workbook wb=new XSSFWorkbook("users.xlsx");
		  
		  Sheet s=wb.getSheetAt(0);
		  
		  
		  FileOutputStream fout=new FileOutputStream("out.xlsx");

		
		try {
		int n=6;
		Row r=s.getRow(1);
		Cell c =r.getCell(0);
	
		
		for(int i=61;i<=88;i++) {
			System.out.println(i);
			r=s.getRow(i);
			c=r.getCell(10);
			double x=c.getNumericCellValue();
			if(x==1) {
				
				c=r.getCell(8);
				c.setCellType(CellType.STRING);
				String id=c.getStringCellValue();
				
					c=r.getCell(5);
					if(c!=null) {
	String url=c.getStringCellValue();
	
				OkHttpClient client = new OkHttpClient().newBuilder()
						  .build();
						MediaType mediaType = MediaType.parse("application/x-www-form-urlencoded");
						RequestBody body = RequestBody.create(mediaType, "avatar="+url);
						Request request = new Request.Builder()
						  .url("https://api-us.cometchat.io/v2.0/users/"+id)
						  .method("PUT", body)
						  .addHeader("Accept", "application/json")
						  .addHeader("Content-Type", "application/x-www-form-urlencoded")
						  .addHeader("appId", "263517b7071bf4c")
						  .addHeader("apiKey", "18902cb47931426ed749d797e910ededa0394b34")
						  .build();
						Response response = client.newCall(request).execute();
						
						String message=response.body().string();
						System.out.println(i+" "+message);
						c=r.createCell(11);
						c.setCellValue(message);
			}
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
