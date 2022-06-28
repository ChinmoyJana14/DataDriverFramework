package Framework.DataDrivenFramework;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConnectExcel {

	public ArrayList<String> getData(String testCaseName) throws IOException {
		// TODO Auto-generated method stub	
		ArrayList<String> data = new ArrayList<String>();
		
		DataFormatter formatter = new DataFormatter();
		
		FileInputStream inputStream = new FileInputStream(System.getProperty("user.dir")+ File.separator +"TestDataFile.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		
		int noOfSheets = workbook.getNumberOfSheets();
		for(int i=0;i<noOfSheets;i++) {
				if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
					XSSFSheet workingSheet = workbook.getSheetAt(i);
					
					Iterator<Row> rows = workingSheet.rowIterator();
					Row firstRow = rows.next();//Select the first row where all header columns are present
						Iterator<Cell> firstRowColumns = firstRow.cellIterator();
							int desiredCol = 0;
							while(firstRowColumns.hasNext()) {
								if(firstRowColumns.next().getStringCellValue().equalsIgnoreCase("TestCase")) {
									// Inside the desired test case column
									System.out.println("Inside the desired test case column");
									System.out.println(desiredCol);	
									break;
								}
								
							}
					while(rows.hasNext()) {
						Row row =rows.next();
						if(row.getCell(desiredCol).getStringCellValue().equalsIgnoreCase(testCaseName)) {
							//Grab all the value of the row    
							Iterator<Cell> testCaseCell = row.cellIterator();
								while(testCaseCell.hasNext()) {
									//System.out.println(testCaseCell.next().getStringCellValue());
									//data.add(testCaseCell.next().getStringCellValue().toString());
									String value = formatter.formatCellValue(testCaseCell.next());
									data.add(value);

								}
						}
					}
				}
		
		}
		return data;
	}
}