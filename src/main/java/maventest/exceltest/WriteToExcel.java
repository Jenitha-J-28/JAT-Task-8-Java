package maventest.exceltest;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToExcel {

	private static WriteToExcel obj1;
	public void WriteToExcelSheet(String sheetName) throws IOException {
		
		FileInputStream fis = new FileInputStream("C:\\Users\\vicky\\eclipse-workspace\\exceltest\\src\\main\\java\\maventest\\exceltest\\EmployeeDatabase.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.createSheet(sheetName);
		
		XSSFRow row = sheet.createRow(0); //0th Row - Column Header
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("Name");
		
		cell = row.createCell(1);
		cell.setCellValue("Age");
		
		cell = row.createCell(2);
		cell.setCellValue("Email Id");
		
		row = sheet.createRow(1); //1st Row
		cell = row.createCell(0);
		cell.setCellValue("John Doe");
		
		cell = row.createCell(1);
		cell.setCellValue("30");
		
		cell = row.createCell(2);
		cell.setCellValue("John@test.com");
		
		row = sheet.createRow(2); //2nd Row
		cell = row.createCell(0);
		cell.setCellValue("Jane Doe");
		
		cell = row.createCell(1);
		cell.setCellValue("28");
		
		cell = row.createCell(2);
		cell.setCellValue("John@test.com");
		
		row = sheet.createRow(3); //3rd Row
		cell = row.createCell(0);
		cell.setCellValue("Bob Smith");
		
		cell = row.createCell(1);
		cell.setCellValue("35");
		
		cell = row.createCell(2);
		cell.setCellValue("Jacky@example.com");
		
		row = sheet.createRow(3); //4th Row
		cell = row.createCell(0);
		cell.setCellValue("Swapnil");
		
		cell = row.createCell(1);
		cell.setCellValue("37");
		
		cell = row.createCell(2);
		cell.setCellValue("Swapnil@example.com");
		
		FileOutputStream fos = new FileOutputStream("C:\\\\Users\\\\vicky\\\\eclipse-workspace\\\\exceltest\\\\src\\\\main\\\\java\\\\maventest\\\\exceltest\\\\EmployeeDatabase.xlsx");
		workbook.write(fos);
		fis.close();
		fos.close();
		workbook.close();
		
	}
	public static void main(String[] args) {
		System.out.println("From the Excel File :");
		WriteToExcel.obj1 = new WriteToExcel();
		try {
			obj1.WriteToExcelSheet("Sheet1");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
				
	}

}
