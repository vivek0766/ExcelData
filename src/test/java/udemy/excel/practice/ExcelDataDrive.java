package udemy.excel.practice;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataDrive {

	
	public ArrayList<String> getData(String TestCaseName) throws IOException {
		
		ArrayList<String> listofTestData =	new ArrayList<String>();	
		FileInputStream inputStream = new FileInputStream("G:\\ApachePOIExample.xlsx");
		XSSFWorkbook xssfworkbook = new XSSFWorkbook(inputStream);
		int noofSheets = xssfworkbook.getNumberOfSheets();
		System.out.println("The No of Sheets Present in Excel file "+noofSheets);
		
		for (int i = 0; i < noofSheets; i++) {

			if (xssfworkbook.getSheetName(i).equalsIgnoreCase("TestCaseData")) {
				System.out.println("The sheet is displyed");
				XSSFSheet sheet = xssfworkbook.getSheetAt(i);
			
				Iterator<Row> rows =sheet.iterator();
				Row firstRow = rows.next();
				Iterator<Cell> firstRowCells =firstRow.cellIterator();
				int index =0;
				int column = 0;
				while(firstRowCells.hasNext()){
					Cell cellValue =firstRowCells.next();
					if(cellValue.getStringCellValue().equalsIgnoreCase("testName")){
						column=index;
					}index++;
					
				}System.out.println(column);
				while(rows.hasNext()){
					Row secondColoumnrows =rows.next();
					if(secondColoumnrows.getCell(column).getStringCellValue().equalsIgnoreCase(TestCaseName)){
						Iterator<Cell> secondcoloumncells=secondColoumnrows.cellIterator();
						while(secondcoloumncells.hasNext()){
						Cell celldata=	secondcoloumncells.next();
						if(celldata.getCellTypeEnum()==CellType.STRING){
						listofTestData.add(celldata.getStringCellValue());
						}else{
							listofTestData.add(NumberToTextConverter.toText(celldata.getNumericCellValue()));
							
						}
						}
					}
				}
			}
		} return listofTestData;
	}
}
