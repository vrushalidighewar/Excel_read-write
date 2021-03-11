package comm;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException {

          XSSFWorkbook workbook=new XSSFWorkbook();
          XSSFSheet sheet=workbook.createSheet("emp");
          
          Map<String, Object[]> data=new TreeMap<String, Object[]>();
          data.put("1",new Object[] {"ID","NAME","LASTNAME"});
          data.put("2",new Object[] {"01","Vrushali","Dighewar"});
          data.put("3",new Object[] {"02","Snehal","Gajbhiye"});
          data.put("4",new Object[] {"03","Indar","Awaghane"});
          
          Set<String> setOFKeys=data.keySet();
          int rownum=0;
          for(String keys: setOFKeys) {
        	  Row row=sheet.createRow(rownum++);
        	  Object[] values=data.get(keys);
          int cellnum=0;
          for(Object obj: values) {
        	  Cell cell=row.createCell(cellnum++);
        	  
              if(obj instanceof String) {
            	  String s=(String) obj;
            	  cell.setCellValue(s);
              }else if(obj instanceof Integer) {	  
        	      Integer i=(Integer) obj;
        	      cell.setCellValue(i);
                   }	  
           }
	}
          FileOutputStream fos=new FileOutputStream("f://Employee.xlsx");
          workbook.write(fos);
          fos.close();
 }
}	
	
