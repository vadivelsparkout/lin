package linkdin;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class data {
	
	@BeforeClass
	void book() throws Exception {
		
	
	  
    //create blank workbook
  XSSFWorkbook workbook=new XSSFWorkbook();
  
  //create blank sheet
  XSSFSheet newsheet=workbook.createSheet("job");
  
  //create row
  XSSFRow row;
  
  //data add
  Map< String, Object[]> job =new TreeMap<String, Object[]>();
  
  job.put("1", new Object[] {"tester"});
  job.put("2", new Object[] {"tester"});
  job.put("3", new Object[] {"tester"});
  job.put("4", new Object[] {"tester"});
  job.put("5", new Object[] {"tester"});
  job.put("6", new Object[] {"tester"});
  job.put("7", new Object[] {"tester"});
  job.put("8", new Object[] {"tester"});
  
  //iterate over data and write to sheet
  Set<String> keyid=job.keySet();
  int rowid=0;
  
  for (String key : keyid) {
	  row=newsheet.createRow(rowid++);
	  Object[] obj=job.get(key);
	  int cellid=0;
	   
	  for (Object obj1 : obj) {
		  Cell cell=row.createCell(cellid++);
		  cell.setCellValue((String) obj1);
		  
		  
		
	}
	
}
  
  //create file
  FileOutputStream out=new FileOutputStream(new File("lin.xlsx"));
  workbook.write(out);
  out.close();
  
}
}
