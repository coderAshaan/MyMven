package ExcelSample;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode {
	
	public static FileInputStream f;
	public static XSSFWorkbook w;
	public static XSSFSheet s;
	
	public static String readStringData(int i , int j) throws IOException { // i for row, j for column
		f = new FileInputStream("C:\\Users\\USER\\Desktop\\ExcelRead\\Recordsone.xlsx");
		w = new XSSFWorkbook(f);
		s = w.getSheet("Sheet1");
		
		Row r = s.getRow(i); // fetching row
		Cell c = r.getCell(j); // fetching column
		
		return c.getStringCellValue();
		
	}
	
	public static double readIntegerData(int i , int j) throws IOException {
		f = new FileInputStream("C:\\Users\\USER\\Desktop\\ExcelRead\\Recordsone.xlsx");
		w = new XSSFWorkbook(f);
		s = w.getSheet("Sheet1");
		
		Row r = s.getRow(i); // fetching row
		Cell c = r.getCell(j); // fetching column
		
		return c.getNumericCellValue();
	}

}
