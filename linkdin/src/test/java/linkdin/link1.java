
	package linkdin;

	import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.WebElement;
	import org.openqa.selenium.chrome.ChromeDriver;
	import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.github.javafaker.Faker;

import java.io.File;
	import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Collector;
import java.util.stream.Collectors;

	public class link1  {
	
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
	  FileOutputStream out=new FileOutputStream(new File("l123.xlsx"));
	  workbook.write(out);
	  out.close();
	  
	}
	

		@Test(priority = 2)
		void launch() throws IOException {
			
			 // Load the xlsx file
            File file = new File("lin.xlsx");
            FileInputStream fis = new FileInputStream(file);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);
            // Get the first row
           
           
            
            
            
			System.setProperty("webdriver.http.factory", "jdk-http-client");

			WebDriver driver = new ChromeDriver();

			ChromeOptions options = new ChromeOptions();
			
		    options.addArguments("-allow-origins", "http://127.0.0.1:45411/");
		    
		    driver.get("https://in.linkedin.com/");
		    
		    WebElement Email = driver.findElement(By.xpath("//input[@id=\"session_key\"]"));
		    Email.sendKeys("7904749236");
		    
		    
		    WebElement Password = driver.findElement(By.xpath("//input[@id=\"session_password\"]"));
		    Password.sendKeys("test@1234");
		    
		    WebElement login =driver.findElement(By.xpath("//button[@type=\"submit\"]"));
		    login.click();
		    
		    
		    List<String> nameList = new ArrayList<String>();
		    
		   Cell cell1;
		    for(int i=0;i<sheet.getLastRowNum();i++) {
		    	cell1=sheet.getRow(i).getCell(0);
		    	WebElement search = driver.findElement(By.xpath("//input[@placeholder=\"Search\"]"));
				   search.sendKeys(cell1.getStringCellValue());
				   search.sendKeys(Keys.ENTER);
				   
				   
				   
				   WebElement people = driver.findElement(By.xpath("//*[@id=\"search-reusables__filters-bar\"]/ul/li[2]"));
		    	   people.click();
		    	
		    	List<WebElement> names=driver.findElements(By.xpath("//ul[@class=\"reusable-search__entity-result-list list-style-none\"]//div[@class=\"t-roman t-sans\"]"));
		    	
		    	
		    	
		    	for (WebElement name : names) {
		    		for(int i1=1;i1<3;i1++) {
		    	    nameList.add(name.getText());

		    	}
		    		}
		    	
		    
		}
		    Map<String, String> map = new HashMap<>();
	        for (int i = 0; i < nameList.size(); i += 2) {
	            String key = nameList.get(i);
	            String value = nameList.get(i + 1);
	            map.put(key, value);
	        }
	        int rowIdx = 0;
	        for (Map.Entry<String, String> entry : map.entrySet()) {
	            Row row1 = sheet.createRow(rowIdx++);
	            Cell keyCell = row1.createCell(1);
	            keyCell.setCellValue(entry.getKey());
	            Cell valueCell = row1.createCell(2);
	            valueCell.setCellValue(entry.getValue());
	        }
	     // Write the workbook to a file
	        FileOutputStream fileOut = new FileOutputStream("l123.xlsx");
	        workbook.write(fileOut);
	        fileOut.close();
			

	
	}
	}

