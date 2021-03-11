package comm;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {
		
         FileInputStream fis=new FileInputStream("f://Employee.xlsx");
         Workbook workbook=new XSSFWorkbook(fis);
         Sheet sheet=workbook.getSheetAt(0);
         Iterator<Row> itRow=sheet.rowIterator();
         while(itRow.hasNext()) {
        	 Row row=itRow.next();
        	 Iterator<Cell> cells=row.cellIterator();
         while(cells.hasNext()) {
        	 Cell cel=cells.next();
         
         if(Cell.CELL_TYPE_NUMERIC == cel.getCellType()) {
        	 System.out.println(cel.getNumericCellValue());
        	 
         }else if(Cell.CELL_TYPE_STRING == cel.getCellType()){
        	 System.out.println(cel.getStringCellValue());
         }
         
         }
         }
	}
}
