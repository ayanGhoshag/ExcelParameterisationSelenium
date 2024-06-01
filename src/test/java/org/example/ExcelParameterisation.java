package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelParameterisation {

    WebDriver driver;

    @BeforeTest
    public void setup() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.get("https://www.saucedemo.com/v1/");
        driver.manage().window().maximize();
    }

    @Test
    public void login() throws IOException, InterruptedException {
        String excelPath = "G:\\ExcelParameterisation\\Book1.xlsx";
        FileInputStream fis = new FileInputStream(excelPath);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet("Sheet1");
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        int rowCount = lastRow - firstRow + 1;
        for (int i = 1; i < rowCount; i++) {
            String username = sheet.getRow(i).getCell(0).getStringCellValue();
            String password = sheet.getRow(i).getCell(1).getStringCellValue();
            driver.findElement(By.xpath("//input[contains(@id,\"user-name\")]")).sendKeys(username);
            driver.findElement(By.xpath("//input[contains(@id,\"password\")]")).sendKeys(password);
            driver.findElement(By.xpath("//input[contains(@id,\"login-button\")]")).click();
            Thread.sleep(2000);
            try {
                if (driver.findElement(By.xpath("//*[@id=\"inventory_filter_container\"]/div")).isDisplayed()) {
                    sheet.getRow(i).createCell(2).setCellValue("sucess");
                    driver.findElement(By.xpath("//button[contains(text(),\"Open Menu\")]")).click();
                    driver.findElement(By.xpath("//a[contains(text(),\"Logout\")]")).click();
                    Thread.sleep(2000);
                }
            } catch (NoSuchElementException e) {
                sheet.getRow(i).createCell(2).setCellValue("Login Failed");
                driver.navigate().refresh();
            }
        }
        FileOutputStream fos = new FileOutputStream(excelPath);
        wb.write(fos);
        wb.close();
    }

    @AfterTest
    public void tearDown() {
        driver.quit();
    }
}
