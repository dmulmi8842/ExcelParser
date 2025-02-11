/**
 * 
 */
package main.webapp.com.excelParser;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;

/**
 * @author Diwash_M Read excel file
 *
 */
public class ReadExcel {
	public static final String COVERSHEETS_FILE_PATH = "D:/Excel documents/GLI Jin Ji Bao Xi - Rising Fortunes_101Z3X.xlsm";

	/**
	 * @param args
	 */
	public static void main(String[] args) throws IOException, InvalidFormatException {
		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook workbook = WorkbookFactory.create(new File(COVERSHEETS_FILE_PATH));

		// Retrieving the number of sheets in the Workbook
		System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

		/*
		 * ============================================================= Iterating over
		 * all the sheets in the workbook (Multiple ways)
		 * =============================================================
		 */

		// Obtain a sheetIterator and iterate over it
		Iterator<Sheet> sheetIterator = workbook.sheetIterator();
		System.out.println("Retrieving Sheets using Iterator");
		while (sheetIterator.hasNext()) {
			Sheet sheet = sheetIterator.next();
			System.out.println("=> " + sheet.getSheetName());
		}

		
		/*
		 * // Or Use for-each loop for(Sheet sheet: workbook) { System.out.println("=> "
		 * + sheet.getSheetName()); }
		 */
		 

		/*
		 * ================================================================== Iterating
		 * over all the rows and columns in a Sheet (Multiple ways)
		 * ==================================================================
		 */

		//Getting sheet with name "Progressive Info"
		Sheet sheet = workbook.getSheet("Progressive Info");
		
		// Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();
        
     // Obtain a rowIterator and columnIterator and iterate over them
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
        	Row row  =  rowIterator.next();
        	
        	// Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }
        
		/*
		 * // 2. Use a for-each loop to iterate over the rows and columns System.out.
		 * println("\n\nIterating over Rows and Columns using for-each loop\n"); for
		 * (Row row: sheet) { for(Cell cell: row) { String cellValue =
		 * dataFormatter.formatCellValue(cell); System.out.print(cellValue + "\t"); }
		 * System.out.println(); }
		 */
        
        workbook.close();
		
	}

}
