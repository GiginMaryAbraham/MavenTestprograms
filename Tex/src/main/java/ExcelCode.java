import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode {

	static FileInputStream f;
	static XSSFWorkbook w;
	static XSSFSheet s;
	static FileOutputStream o;
	static XSSFWorkbook w1;
	static XSSFSheet s1;

	public static String getStringData(int a, int b) throws IOException {

		f = new FileInputStream("\\D:\\testexcel.xlsx");
		w = new XSSFWorkbook(f);
		s = w.getSheet("Sheet1");
		Row r = s.getRow(a);
		Cell c = r.getCell(b);
		return c.getStringCellValue();
	}

	public static String getIntegerData(int a, int b) throws IOException {

		f = new FileInputStream("\\D:\\testexcel.xlsx");
		w = new XSSFWorkbook(f);
		s = w.getSheet("Sheet1");
		Row r = s.getRow(a);
		Cell c = r.getCell(b);
        int x = (int) c.getNumericCellValue();
		return String.valueOf(x);

	}

	
	  public static void setData() throws IOException {
	  
	  o = new FileOutputStream("\\D:\\testexcel.xlsx"); 
	  w1 = new XSSFWorkbook(); 
	  s1 = w1.createSheet("Sheet2"); 
	  Row r = s1.createRow(0); 
	  Cell c =r.createCell(0); 
	  c.setCellValue("Employee");
	  w1.write(o);
	  
	  }
	 

}
