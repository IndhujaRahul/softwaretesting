package org.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {

	public static WebDriver driver;
	
	public static void launchBrowser()
	{
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		
	}
	public static void windowMaximize()
	{
		driver.manage().window().maximize();
	}
	public static void launchUrl(String url)
	{
		driver.get(url);
	}
	public static void pageUrl()
	{
		String url = driver.getPageSource();
		System.out.println(url);
	}
	public static void passText(String txt, WebElement ele)
	{
		ele.sendKeys(txt);
	}
	public static void closeentireBrowser()
	{
		driver.quit();
	}
	public static void clickBtn(WebElement ele)
	{
		ele.click();
	}
	public static void screenShot(String imgName) throws IOException
	{
		TakesScreenshot ts = (TakesScreenshot) driver;
		File image = ts.getScreenshotAs(OutputType.FILE);
		File f = new File("location+ imgName.png");
		FileUtils.copyFile(image, f);
	}
	public static Actions a;
	
	public static void moveThecursor(WebElement targetwebelement)
	{
		a = new Actions(driver);
		a.moveToElement(targetwebelement).perform();
	}
	public static void dragDrop(WebElement dragWebelement, WebElement dragElemnt)
	{
		a = new Actions(driver);
		a.dragAndDrop(dragWebelement, dragElemnt);
	}
	public static JavascriptExecutor js;
	public static void scrollThePage(WebElement tarWebElement)
	{	
		js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].scrollIntoView(true)", tarWebElement);
	}
	public static void scroll(WebElement element)
	{
		js =(JavascriptExecutor) driver;
		js.executeScript("arguments[0].scrollIntoView(false)", element);
	}
	public static void excelRead(String sheetName, int row, int cell) throws IOException 
	{
		File f = new File("excellocation.xlsx");
	    FileInputStream fls = new FileInputStream(f);
	    Workbook wb = new XSSFWorkbook(fls);
	    Sheet mysheet = wb.getSheet("Data");
	    Row r = mysheet.getRow(row);
	    Cell c = r.getCell(cell);
	    int cellType = c.getCellType();
	    String Value = " ";
	    
	    if (cellType==1) {
	    	String value2 = c.getStringCellValue();
			
		}
	    else if (DateUtil.isCellDateFormatted(c)) {
	    	Date  dt = c.getDateCellValue();
	    	SimpleDateFormat s = new SimpleDateFormat(Value);
	    	String value1 = s.format(dt);
	    	
			
		}
	    else {
			double d = c.getNumericCellValue();
			long l = (long) d;
			String valueOf = String.valueOf(l);
		}
	
	    }
	public static void CreateNewExcelFile(int getRow, int getCell, String newData) throws IOException
	{
		File f = new File("Excellocation.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet newSheet = w.createSheet("Datas");
		Row r = newSheet.createRow(getRow);
		Cell c = r.createCell(getCell);
		c.setCellValue(newData);
		FileOutputStream fos = new FileOutputStream(f);
		w.write(fos);
		
	}
	public static void createCell(int getRow, int getCell, String newData) throws IOException
	{
		File f = new File("Excellocation.xlsx");
		FileInputStream fls = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fls);
		Sheet s = w.getSheet("Datas");
		Row r = s.getRow(getRow);
		Cell cl = r.createCell(getCell);
		cl.setCellValue(newData);
		FileOutputStream fos = new FileOutputStream(f);
		w.write(fos);
		}
	public static void createRow(int cRow, int cCell, String newData) throws IOException
	{
		File f = new File("Excellocation.xlsx");
		FileInputStream fls = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fls);
		Sheet s = w.getSheet("Datas");
		Row r = s.createRow(cRow);
		Cell c = r.createCell(cCell);
		c.setCellValue(newData);
		FileOutputStream fos = new FileOutputStream(f);
		w.write(fos);
		
	}
	public static void updateData(int getRow, int getCell, String ExistingData, String newData) throws IOException
	{
		File f = new File("Excel.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fis);
		Sheet s = w.getSheet("Data");
		Row r = s.getRow(getRow);
		Cell c = r.getCell(getCell);
		String str = c.getStringCellValue();
		if (str.equals(ExistingData)) {
			c.setCellValue(newData);
		}
		FileOutputStream fos = new FileOutputStream(f);
		w.write(fos);
	}
}
