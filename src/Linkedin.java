import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Linkedin {

    public static void main(String[] args) throws AWTException, IOException {
        System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        FileOutputStream fos = new FileOutputStream("test5.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("list");
        int count = 1;
        Row row = sheet.createRow(count);
        Cell c = row.createCell(0);

        try {
            driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
            driver.get("https://www.linkedin.com/search/results/people/?keywords=fintech%20partnerships%2C%20fintech%2C%20india&origin=GLOBAL_SEARCH_HEADER");
            driver.findElement(By.className("main__sign-in-link")).click();
            Thread.sleep(2);
            driver.findElement(By.id("username")).sendKeys("dhirajuchil51@gmail.com");
            driver.findElement(By.id("password")).sendKeys("zynga123@");

            Robot r = new Robot();
            r.keyPress(KeyEvent.VK_ENTER);

            // Create instance of Javascript executor

            for (int page = 1; page < 100; page++) {
                for (int i = 1; i <= 10; i++) {

                    row = sheet.createRow(count++);


//		WebElement ele=driver.findElement(By.xpath("//*[@id=\"ember61\"]/div/ul/li["+i+"]/div/div/div[2]/a"));
                    WebElement ele = driver.findElement(By.xpath("/html/body/div[8]/div[3]/div/div[1]/div/div[1]/main/div/div/div[2]/ul/li[" + i + "]/div/div/div[2]/div[1]/div/div[1]/span/div/span[1]/span/a"));

                    if (i == 4) {
                        r.keyPress(KeyEvent.VK_PAGE_DOWN);
                        Thread.sleep(500);
                        r.keyPress(KeyEvent.VK_PAGE_DOWN);
                    }
                    String url = ele.getAttribute("href");
                    String name = "";
                    try {
                        name = driver.findElement(By.xpath("/html/body/div[8]/div[3]/div/div[1]/div/div[1]/main/div/div/div[2]/ul/li[" + i + "]/div/div/div[2]/div[1]/div/div[1]/span/div/span[1]/span/a/span/span[1]")).getText();
                        System.out.println(name);
                    } catch (Exception e) {
                        System.out.println("linkdin member");
                        continue;
                    }
//			String name=driver.findElement(By.xpath("//*[@class=\"pv2 artdeco-card ph0 mb2\"]/div/ul/li["+i+"]/div/div/div[2]/a/h3/span/span")).getText();
                    String designation = "";
                    String company = "";
                    try {
                        //ele.getAttribute("href");
                        String details = driver.findElement(By.xpath("/html/body/div[8]/div[3]/div/div[1]/div/div[1]/main/div/div/div[2]/ul/li[" + i + "]/div/div/div[2]/div[1]/div/div[2]/div[1]")).getText();
//		String details=driver.findElement(By.xpath("//*[@id=\"ember61\"]/div/ul/li["+i+"]/div/div/div[2]/p")).getText();

                        String d[] = details.split(" at ");
                        designation = d[0];

                        company = d.length >= 2 ? d[1] : "";
                        System.out.println(designation + " " + company);
                    } catch (Exception e) {
                        System.out.println("linkdin member");
                        continue;
                    }
                    c = row.createCell(0);
                    c.setCellValue(name.split("\n")[0]);
                    c = row.createCell(1);
                    c.setCellValue(designation);
                    c = row.createCell(2);
                    c.setCellValue(company);

                    c = row.createCell(3);
                    c.setCellValue(url);
                }
                System.out.println("---------------------" + (page + 1) + "------------------------------");
                driver.get("https://www.linkedin.com/search/results/people/?keywords=fintech%20partnerships%2C%20fintech%2C%20india&origin=GLOBAL_SEARCH_HEADER&page="+(page+1));
            }
        } catch (Exception e) {
            // TODO: handle exception
            e.printStackTrace();
        }

        workbook.write(fos);
        fos.flush();
        fos.close();

    }

}
