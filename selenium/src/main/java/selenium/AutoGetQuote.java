package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class AutoGetQuote {

    String now;
    
    static int COL_COMPANY = 0;
    static int COL_CODE = 1;
    static int COL_NUM_HOLDINGS = 2;
    static int COL_DATE = 3;
    static int COL_PRICE = 4;

    public AutoGetQuote() {
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy-HH:mm");  
        now = formatter.format(new Date());  
    }
    
    static double getQuote(String code) throws Exception {

        System.setProperty("webdriver.chrome.driver", "D:\\Software\\chromedriver_85.exe");

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");
        WebDriver driver = new ChromeDriver(options);

        driver.get("https://finance.yahoo.com/quote/" + code);
        Thread.sleep(5000);

        WebElement we = driver.findElement(By.cssSelector("span[data-reactid=\"32\"]"));

        String price = we.getText();

        driver.close();

        return NumberFormat.getInstance().parse(price).doubleValue();
    }

    public void getNetworthForExcel(String file) throws Exception {
        
        FileInputStream fis = new FileInputStream(new File(file));
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
            
            double price = getQuote(code);
            System.out.println(code + " " + price);
        
            row.getCell(COL_DATE).setCellValue(now);
            row.getCell(COL_PRICE).setCellValue(price);
        }
        
        fis.close();
        
        FileOutputStream fos = new FileOutputStream(new File(file));
        wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
        wb.write(fos);
        fos.close();
    }
    
    public static void main(String[] args) throws Exception {
    	
//    	System.out.println(By.cssSelector("span[data-reactid=\"32\"]"));
        AutoGetQuote autoGetQuote = new AutoGetQuote();
        
        autoGetQuote.getNetworthForExcel("D:\\holdings.xlsx");
    }
}
