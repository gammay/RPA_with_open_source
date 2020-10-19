package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

//import java.awt.Toolkit;
//import java.awt.datatransfer.Clipboard;
//import java.awt.datatransfer.DataFlavor;

public class AutoGetQuote1 {

	String now;
	
	static int COL_COMPANY = 0;
	static int COL_CODE = 1;
	static int COL_NUM_HOLDINGS = 2;
	static int COL_DATE = 3;
	static int COL_PRICE = 4;

	public AutoGetQuote1() {
	    SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy-HH:mm");  
	    now = formatter.format(new Date());  
	}
	
	static String getQuote(String code) throws Exception {

		System.setProperty("webdriver.chrome.driver", "D:\\Software\\chromedriver_85.exe");

		ChromeOptions options = new ChromeOptions();
		options.addArguments("--start-maximized");
		WebDriver driver = new ChromeDriver(options);

		driver.get("https://finance.yahoo.com/quote/" + code);
		Thread.sleep(5000);

		WebElement we = driver.findElement(By.cssSelector("span[data-reactid=\"32\"]"));
//		Thread.sleep(5000);

		String price = we.getText();

		driver.close();

//		Actions actions = new Actions(driver);
//		actions.moveToElement(we, 10, 10).doubleClick().build().perform();
//		Thread.sleep(2000);
//		we.sendKeys(Keys.CONTROL, "c");
//		Toolkit toolkit = Toolkit.getDefaultToolkit();
//		Clipboard clipboard = toolkit.getSystemClipboard();
//		String result = (String) clipboard.getData(DataFlavor.stringFlavor);
//		System.out.println("String from Clipboard:" + result);

		System.out.println("*** " + price);
		return price;
	}

//	void getSymbols() throws Exception {
//		FileInputStream fis = new FileInputStream(new File("D:\\holdings.xlsx"));
//		XSSFWorkbook wb = new XSSFWorkbook(fis);
//		XSSFSheet sheet = wb.getSheetAt(0);
//		FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
//		for (Row row : sheet) // iteration over row using for each loop
//		{
//			String company = row.getCell(COL_COMPANY).getStringCellValue();
//			if(company.equals("COMPANY")) {
//				continue;
//			}
//			if(company.equals("NETWORTH")) {
//				break;
//			}
//			String code = row.getCell(COL_CODE).getStringCellValue();
//			int num_holding = (int)row.getCell(COL_NUM_HOLDINGS).getNumericCellValue();
//			
//			companies.add(company);
//			symbols.add(code);
//			num_holdings.add(num_holding);
//		}
//	}
	
	public void getNetworth() throws Exception {
		
		FileInputStream fis = new FileInputStream(new File("D:\\holdings.xlsx"));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);

		for (Row row : sheet)
		{
			String company = row.getCell(COL_COMPANY).getStringCellValue();
			if(company.equals("COMPANY")) {
				continue;
			}
			if(company.equals("NETWORTH")) {
				break;
			}
			String code = row.getCell(COL_CODE).getStringCellValue();
			
			String price = getQuote(code);
//			String price = "1,234.56";
			System.out.println("+++ " + code + " " + price);
		
            row.getCell(COL_DATE).setCellValue(now);
            row.getCell(COL_PRICE).setCellValue(price);
		}
		
		fis.close();
		
		FileOutputStream fos = new FileOutputStream(new File("D:\\holdings.xlsx"));
		wb.write(fos);
        fos.close();
	}
	
	public static void main(String[] args) throws Exception {
//		GetQuote.getQuote("GOOG");
		
		AutoGetQuote1 autoGetQuote = new AutoGetQuote1();
		
		System.out.println(autoGetQuote);
		
//		autoGetQuote.getSymbols();
		autoGetQuote.getNetworth();
		
//		System.out.println(autoGetQuote);
	}
}
