import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class Nagwa {
    public static void main(String[] args) throws IOException {

        //open chrome and Navigate to the URL
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Ahmed\\Desktop\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.nagwa.com/");
        driver.manage().window().maximize();
        //select the language
        WebElement language  = driver.findElement(By.xpath("/html/body/div/div/main/div[2]/ul/li[1]/a"));
        language.click();
        //Search for the required item
        driver.manage().timeouts().implicitlyWait(4, TimeUnit.SECONDS);
        WebElement search = driver.findElement(By.xpath("/html/body/header/div[1]/div/div[3]/button"));
        search.click();
        //Read the search Keyword from Extrnal excl Sheet
        String data = ExcelOperations.importDataFromExcel();
        driver.findElement(By.xpath("//*[@id=\"txt_search_query\"]")).sendKeys(data);
        WebElement searchBtn  = driver.findElement(By.id("btn_global_search"));
        searchBtn.click();
        WebElement lesson   = driver.findElement(By.xpath("/html/body/div/div[1]/div/div/main/div[3]/ul/li[2]/div/a"));
        lesson .click();
        WebElement workSheet   = driver.findElement(By.xpath("//*[@id=\"WorksheetSection\"]/div/div[2]/div/a"));
        workSheet .click();
        //count the questions
        driver.manage().timeouts().implicitlyWait(4, TimeUnit.SECONDS);
        List<WebElement> questionsCount = driver.findElements(By.className("instance"));
        int numberofquestions = questionsCount.size() ;
        System.out.println("Total number of questions is " + numberofquestions);
        driver.close();
        // print the count at Extrnal sheet
        ExcelOperations.writeToExcel(numberofquestions);
    }
}











