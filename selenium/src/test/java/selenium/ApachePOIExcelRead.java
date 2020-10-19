package selenium;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ApachePOIExcelRead {
	public static void main(String args[]) throws Exception {
		//Load the workbook into the memory
		FileInputStream fis = new FileInputStream("D://holdings.xlsx");
		Workbook workbook = WorkbookFactory.create(fis);

		//Modify the workbook as you wish
		//As an example, we override the first cell of the first row in the first sheet (0-based indices)
		workbook.getSheetAt(0).getRow(0).getCell(0).setCellValue("new value for A1");


		//you have to close the input stream FIRST before writing to the same file.
		fis.close() ;

		//save your changes to the same file.
		workbook.write(new FileOutputStream("D://holdings.xlsx")); 
		workbook.close();
		
////obtaining input bytes from a file  	
//		FileInputStream fis = new FileInputStream(new File("D:\\holdings.xlsx"));
////creating workbook instance that refers to .xls file  
//		XSSFWorkbook wb = new XSSFWorkbook(fis);
////creating a Sheet object to retrieve the object  
//		XSSFSheet sheet = wb.getSheetAt(0);
////evaluating cell type   
//		FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
//		for (Row row : sheet) // iteration over row using for each loop
//		{
//			for (Cell cell : row) // iteration over cell using for each loop
//			{
//				switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
//				case Cell.CELL_TYPE_NUMERIC: // field that represents numeric cell type
////getting the value of the cell as a number  
//					System.out.println(cell.getNumericCellValue());
//					break;
//				case Cell.CELL_TYPE_STRING: // field that represents string cell type
////getting the value of the cell as a string  
//					System.out.println(cell.getStringCellValue());
//					break;
//				}
//			}
//			System.out.println();
//		}
	}
}

//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.IOException;
//import java.util.Iterator;
//
//public class ApachePOIExcelRead {
//
//    private static final String FILE_NAME = "D:\\holdings.xlsx";
//
//    public static void main(String[] args) {
//
//        try {
//
//            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
//            Workbook workbook = new XSSFWorkbook(excelFile);
//            Sheet datatypeSheet = workbook.getSheetAt(0);
//            Iterator<Row> iterator = datatypeSheet.iterator();
//
//            while (iterator.hasNext()) {
//
//                Row currentRow = iterator.next();
//                Iterator<Cell> cellIterator = currentRow.iterator();
//
//                while (cellIterator.hasNext()) {
//
//                    Cell currentCell = cellIterator.next();
//                    //getCellTypeEnum shown as deprecated for version 3.15
//                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
//                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
//                        System.out.print(currentCell.getStringCellValue() + "--");
//                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
//                        System.out.print(currentCell.getNumericCellValue() + "--");
//                    }
//
//                }
//                System.out.println();
//
//            }
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//    }
//}