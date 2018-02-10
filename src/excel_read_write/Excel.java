package excel_read_write;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Excel {
	        String key="www.chrome.driver";
			String value="G:\\SELENIUM\\chromedriver.exe";
	        WebDriver driver;  
	        XSSFWorkbook Workbook;
	        XSSFSheet sheet;
	        XSSFCell cell;
    @BeforeTest
	public void setup(){
		try {
		System.setProperty(key, value);
		driver=new ChromeDriver();
		driver.navigate().to("https://www.google.com");
		driver.manage().window().maximize();
		driver.navigate().to("https://www.amazon.in/ap/register?openid.pape.max_auth_age=0&openid."
				+ "identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&pageId="
				+ "inflex&ignoreAuthState=1&openid.return_to="
				+ "https%3A%2F%2Fwww.amazon.in%2F%3Fref_%3Dnav_ya_signin&prevRID=68C7QRDXYG5DH9Q1"
				+ "E4R3&openid.assoc_handle=inflex&openid.mode=checkid_setup&openid.ns.pape=http%3A%2F%2Fspecs.openid.net%2Fextensions%2Fpape%2F1.0&prepopulatedLoginId=&failedSignInCount=0&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns="
				+ "http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0");
		
		
		
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		}
    
		@Test
		public void read_data(){
			try {
				File src=new File("G:\\SELENIUM\\Read_Write Excel file\\amazon.xlsx");
				FileInputStream fis=new FileInputStream(src);
				Workbook = new XSSFWorkbook(fis);
				sheet = Workbook.getSheetAt(0);
				int cellnum=sheet.getLastRowNum();
				for(int i=1;i<=cellnum;i++){
					 cell = sheet.getRow(i).getCell(0);
					 cell.setCellType(cell.CELL_TYPE_STRING);
					  driver.findElement(By.xpath("//*[@id='ap_customer_name']")).sendKeys(cell.getStringCellValue());
		 
		 driver.findElement(By.cssSelector("#auth-country-picker-container > span > span > span")).click();
		 driver.findElement(By.cssSelector("#auth-country-picker_90")).click();
		 
	            cell =sheet.getRow(i).getCell(1);
	            cell.setCellType(cell.CELL_TYPE_STRING);
	            driver.findElement(By.cssSelector("#ap_phone_number")).sendKeys(cell.getStringCellValue());
	            
	            cell=sheet.getRow(i).getCell(2);
	            cell.setCellType(cell.CELL_TYPE_STRING);
	            driver.findElement(By.cssSelector("#ap_email")).sendKeys(cell.getStringCellValue());
	            
	            cell=sheet.getRow(i).getCell(3);
	            cell.setCellType(cell.CELL_TYPE_STRING);
	            driver.findElement(By.cssSelector("#ap_password")).sendKeys(cell.getStringCellValue());
	            
	            driver.findElement(By.cssSelector("#ap_register_form > div > div > div:nth-child(9) > a")).click();
	            
	           
	            
					
				}
				
				
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
			}
			
			
			
		}
		
		
		
		
		
		
		
		
	}
		
	


