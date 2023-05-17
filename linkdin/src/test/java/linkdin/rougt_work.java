package linkdin;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class rougt_work {

	@Test
	void lin() throws IOException {
		// create blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// create blank sheet
		XSSFSheet newsheet = workbook.createSheet("job");

		// create row
		XSSFRow row;

		// data add
		List<String> nameList = new ArrayList<String>();
		nameList.add("one");
		nameList.add("two");
		nameList.add("three");
		nameList.add("four");
		nameList.add("five");
		nameList.add("six");

		Map<String, String> map = new HashMap<>();
        for (int i = 0; i < nameList.size(); i += 2) {
            String key = nameList.get(i);
            String value = nameList.get(i + 1);
            map.put(key, value);
        }
        int rowIdx = 0;
        for (Map.Entry<String, String> entry : map.entrySet()) {
            Row row1 = newsheet.createRow(rowIdx++);
            Cell keyCell = row1.createCell(1);
            keyCell.setCellValue(entry.getKey());
            Cell valueCell = row1.createCell(2);
            valueCell.setCellValue(entry.getValue());
        }
     // Write the workbook to a file
        FileOutputStream fileOut = new FileOutputStream("n.xlsx");
        workbook.write(fileOut);
        fileOut.close();
		}
	@Test
	void nex() throws IOException {
	
			// create blank workbook
			XSSFWorkbook workbook = new XSSFWorkbook();

			// create blank sheet
			XSSFSheet newsheet = workbook.createSheet("job");

			// create row
			XSSFRow row;
			
		List<String> nameList = new ArrayList<String>();
		nameList.add("one");
		nameList.add("two");
		nameList.add("three");
		nameList.add("four");
		nameList.add("five");
		nameList.add("six");
		
		Row dataRow = newsheet.createRow(0);

		for (int i = 0; i < nameList.size(); i++) {

			

			dataRow.createCell(i+1).setCellValue(nameList.get(i)); // Set value in the second column (column index 1)

			// Write the workbook to a file
			FileOutputStream fileOut = new FileOutputStream("nm.xlsx");
			
			workbook.write(fileOut);
			
			fileOut.close();

		}



	}

	}


