import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Scrap {
	public static void main(String[] args) throws IOException, InterruptedException {
		Workbook wb=new XSSFWorkbook("file.xlsx");
		
		Sheet s=wb.getSheetAt(0);
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("https://finovate.com/videos/?filtertype=&showtypes=Show&videostartyear=2019&showletters=A-Z");
		
		driver.findElement(By.id("cn-accept-cookie")).click();
		
		/*
		 * WebElement html = driver.findElement(By.tagName("html"));
		 * html.sendKeys(Keys.chord(Keys.CONTROL, Keys.SUBTRACT));
		 * html.sendKeys(Keys.chord(Keys.CONTROL, Keys.SUBTRACT));
		 * html.sendKeys(Keys.chord(Keys.CONTROL, Keys.SUBTRACT));
		 * html.sendKeys(Keys.chord(Keys.CONTROL, Keys.SUBTRACT));
		 */
		
		int lineNo=1;

		try {
			for(int p=1;p<=9;p++) {
		for(int i=1;i<=27;i++) {
			
			if(p==1&&i==27)
				continue;
			
			if(i==10) {		
				JavascriptExecutor js = (JavascriptExecutor) driver;
		        js.executeScript("javascript:window.scrollBy(350,500)");
		        Thread.sleep(1000);
			}
				
		Row r=s.createRow(lineNo++);
		int j=0;
		Cell c=r.createCell(j++);
		
		
		String name=driver.findElement(By.xpath("//*[@id=\"main\"]/div/div[2]/div["+i+"]/p[1]/a")).getText();
		
		//((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", driver.findElement(By.xpath("//*[@id=\"main\"]/div/div[2]/div["+i+"]/p[1]/a")));
		
	
		
		
		
		c.setCellValue(name);
	//	c=r.createCell(j++);
		
		
		
		System.out.println(name);
		driver.findElement(By.xpath("//*[@id=\"main\"]/div/div[2]/div["+i+"]/p[1]/a")).click();
		Thread.sleep(1000);
		
		String info=driver.findElement(By.xpath("//*[@id=\"main\"]/div/div[2]/div[2]/p[2]")).getText();
				/*
				 * String a[]=info.split("\n"); for(String line:a) { c=r.createCell(j++);
				 * c.setCellValue(line); }
				 */
		
		c=r.createCell(j++);
		c.setCellValue(info);
		
		String info2=driver.findElement(By.xpath("//*[@id=\"main\"]/div/div[2]/div[2]/p[3]")).getText();
				
				  String b[]=info2.split("\n"); for(String line:b) { c=r.createCell(j++);
				  c.setCellValue(line); }
				  
				  if(b.length<5) {
					  j+=5-b.length;
				  }
				 
		
					/*
					 * c=r.createCell(j++); c.setCellValue(info2);
					 */
		
		List<WebElement> lol=driver.findElements(By.xpath(("//*[@id=\"main\"]/div/div[2]/div[1]/p")));
		System.out.println(lol.size());
		
		int k=lol.size();
		String contacts=driver.findElement(By.xpath(("//*[@id=\"main\"]/div/div[2]/div[1]/p["+k+"]"))).getText();
		String x[]=contacts.split("\n");
				
				 for(String line:x) { c=r.createCell(j++); c.setCellValue(line);
				 
				 System.out.println(line);}
				 
		driver.navigate().back();
		//Thread.sleep(500);
		}
		
				/*
				 * JavascriptExecutor js = (JavascriptExecutor) driver;
				 * js.executeScript("javascript:window.scrollBy(350,500)"); Thread.sleep(1000);
				 * 
				 * driver.findElement(By.xpath("//*[@id=\"main\"]/div/div[3]/div/nav/ul/li["+(p+
				 * 2)+"]/a")).click();
				 */
		driver.get("https://finovate.com/videos/page/"+(p+1)+"/?filtertype&showtypes=Show&videostartyear=2019&showletters=A-Z");
		
			}
		}
		catch(Exception e) {
			e.printStackTrace();
		}
		
		
		
		
		FileOutputStream fout=new FileOutputStream("out.xlsx");
		wb.write(fout);
		fout.close();
		wb.close();
		}
		
		
		
		
		
		
	

}
