package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	String now;
	
	static int COL_COMPANY = 0;
	static int COL_CODE = 1;
	static int COL_NUM_HOLDINGS = 2;
	static int COL_DATE = 3;
	static int COL_PRICE = 4;

	public static void main(String[] args) throws Exception {
		FileInputStream fis = new FileInputStream(new File("D:\\holdings.xlsx"));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);

		FileOutputStream fos = new FileOutputStream(new File("D:\\holdings.xlsx"));
		for (Row row : sheet)
		{
			row.getCell(0).setCellValue("xxx");
		}
		
		fis.close();
		wb.write(fos);
        fos.close();		
	}
}
