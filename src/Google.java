import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
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

public class Google {

    public static void main(String[] args) throws IOException {

        // TODO Auto-generated method stub
        Workbook wb = new XSSFWorkbook("file.xlsx");

        Sheet s = wb.getSheetAt(0);


        FileOutputStream fout = new FileOutputStream("out.xlsx");

        System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
        DesiredCapabilities capabilities = DesiredCapabilities.chrome();
        capabilities.setCapability("chrome.switches", Arrays.asList("--incognito"));
        ChromeOptions options = new ChromeOptions();
        options.addArguments("headless");
        try {
            int n = 6;
            Row r = s.getRow(1);
            Cell c = r.getCell(0);
            WebDriver driver = new ChromeDriver();

            for (int i =1001 ; i <= 1500; i++) {

                try {
                    r = s.getRow(i);

                    String st = "";
                    c = r.getCell(0);
                    st += c.getStringCellValue();
                    c = r.getCell(1);
                    st += "+" + c.getStringCellValue();
                    c = r.getCell(2);
                    st += "+" + c.getStringCellValue();
                    st += "+linkedin";

                    //driver = new ChromeDriver(options);
                    //driver.manage().window().setPosition(new Point(-2000, 0));
                    //	driver = new ChromeDriver();
                    driver.get("https://www.google.com/search?q=" + st);

                    driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
                    List<WebElement> w = driver.findElements(By.cssSelector("div.g"));
                    WebElement link = w.get(0).findElement(By.tagName("a"));
                    String l = link.getAttribute("href");
                    System.out.println(i + " url:" + l);

                    c = r.createCell(6);
                    c.setCellValue(l);
                    //	driver.close();
                } catch (Exception e) {
                    e.printStackTrace();
                       driver.close();
                     driver = new ChromeDriver(capabilities);


                }


            }
            driver.close();
        } catch (Exception e) {
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
