package dataDriven.excelDataProvider;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.io.FileInputStream;
import java.io.IOException;

public class DataProvide {

	@Test(dataProvider = "driveTest")
	public void testCaseData(String greeting,String communication,String id) {
		System.out.println(greeting+communication+id);
	}
	
	@DataProvider(name="driveTest")
	public Object[][]  getData() throws IOException {

		FileInputStream fis =  new FileInputStream("C:\\Selenium\\message.xlsx");
		XSSFWorkbook workBook =  new XSSFWorkbook(fis);
		XSSFSheet  sheet = workBook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int noOfColumn = row.getLastCellNum();
		Object[][]  data = new Object[rowCount-1][noOfColumn];

		for (int i = 0; i < rowCount-1; i++) {
			 row = sheet.getRow(i);
			for (int j = 0; j < noOfColumn; j++) {
				
				data[i][j]= new DataFormatter().formatCellValue(row.getCell(j));
			
		}
		}
		return data;

	}
}
