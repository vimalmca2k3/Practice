
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSReading {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream ("C:\\util\\xls\\TestSuite.xlsx");
		
	
		XSSFWorkbook  testSuite	= new XSSFWorkbook (fis);
		XSSFSheet TestCases = testSuite.getSheetAt(0);
		
		XSSFRow row ;
		int noOfRows = TestCases.getLastRowNum();
		
		
		for (int i =1 ; i < noOfRows+1 ; i++)
			
		{
			row = TestCases.getRow(i);
			
			for (int j=0 ; j< row.getLastCellNum() ; j++)
				
			{
				
				System.out.print(row.getCell(j).getStringCellValue());
				System.out.println("         ");
							
			}
			
			System.out.println();
			
		}
		
		
		

	}

	
}
