package maventest.exceltest;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcel {

	public void ReadFromExcelSheet(String sheetName) throws IOException {
		FileInputStream fis = new FileInputStream("C:\\Users\\vicky\\eclipse-workspace\\exceltest\\src\\main\\java\\maventest\\exceltest\\EmployeeDatabase.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		
		int lastRow = sheet.getLastRowNum();
		for (int i=0; i<=lastRow; i++) {
			XSSFRow row = sheet.getRow(i);
		int lastColumn = row.getLastCellNum();
		for (int j=0; j<lastColumn; j++) {
			XSSFCell cell = row.getCell(j);
			String result = cell.getStringCellValue();
			System.out.print(result + "  ");
		}
			System.out.println();
		}
	}
	public static void main(String[] args) {
		ReadFromExcel read = new ReadFromExcel();
		try {
			read.ReadFromExcelSheet("Sheet1");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}

}
