package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class DataFromExcelRunner {

	@Test
	public void getFromExcel() throws IOException {
		FileInputStream fis = new FileInputStream("C:\\Selenium\\message.xlsx");
		XSSFWorkbook workBook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workBook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int noOfColumn = row.getLastCellNum();
		Object[][] data = new Object[rowCount - 1][noOfColumn];

		for (int i = 0; i < rowCount-1; i++) {
			 row = sheet.getRow(i);
			for (int j = 0; j < noOfColumn; j++) {
				
				data[i][j]= new DataFormatter().formatCellValue(row.getCell(j));
			
		}
		}

	}
}
