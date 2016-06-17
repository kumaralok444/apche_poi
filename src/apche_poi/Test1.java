package apche_poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Test1 {
	WebDriver driver;
	
	
	
	@DataProvider
	public Object[][] loginData() {
		Object[][] arrayObject = getExcelData("D:/Alok/data.xlsx","Sheet1");
		return arrayObject;
	}
	
	
	
	@BeforeTest
	public void setUp()
	{
		driver=new FirefoxDriver();
		driver.get("file:///C:/Users/NEW/Desktop/test.html");
		driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
		driver.manage().window().maximize();
	}
	
	
	
	@Test(dataProvider="loginData")
	public void testCAse1(String ce1,String ce2)
	{  
		WebElement wb;
		wb=driver.findElement(By.id("userName"));
		wb.clear();
		wb.sendKeys(ce1);
		wb=driver.findElement(By.id("password"));
		wb.clear();
		wb.sendKeys(ce2);
		wb=driver.findElement(By.name("login"));
		wb.click();
		//Alert a=driver.switchTo().alert();
		//a.accept();
	}
	
	@AfterTest
	public void tearDown()
	{
		driver.quit();
	}
	
	public String[][] getExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = new XSSFWorkbook(fs);
			Sheet sh = wb.getSheet(sheetName);
			Row r=sh.getRow(0);

			int totalNoOfCols = r.getLastCellNum();
			int totalNoOfRows = sh.getLastRowNum();
			//int r1=sh.getFirstRowNum();
			//System.out.println("Total num of column:"+totalNoOfCols);
			//System.out.println("Total num of Rows:"+totalNoOfRows);
			//System.out.println("First Row num is:"+r1);
			arrayExcelData = new String[totalNoOfRows][totalNoOfCols];
			
			for (int i= 1 ; i <=totalNoOfRows; i++) {
				r=sh.getRow(i);
				for (int j=0; j < totalNoOfCols; j++) {
					arrayExcelData[i-1][j] = r.getCell(j).toString();
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} 
		return arrayExcelData;
	}
}
