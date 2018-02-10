package utility;

import java.io.File;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class test_data {
	static File file;
	static Workbook data;
	static Sheet sheet;
	
	public static Object[][] getAmazonData(String amazon){
		try {
			file=new File("G:\\SELENIUM\\Read_Write Excel file\\src\\testdata\\Excel_selenium.xlsx");
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		try {
			data = WorkbookFactory.create(file);
		} catch (Exception e) {
			e.printStackTrace();
		}
		sheet = data.getSheet(amazon);
		
		Object[][] impdata = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		
		for(int i=0;i<sheet.getLastRowNum();i++){
			for(int k=0;k<sheet.getRow(0).getLastCellNum();k++){
				impdata[i][k]=sheet.getRow(i+1).getCell(k);
				
			}
			
		}
		return impdata;
	}

}
