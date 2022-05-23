package week5.day2;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public static String[][] excel() throws IOException {
	    XSSFWorkbook book=new XSSFWorkbook("./Data/TestNGInputExcel.xlsx");
	    XSSFSheet sheet = book.getSheet("Sheet1");
	    int lastRowNum = sheet.getLastRowNum();
	    System.out.println(lastRowNum);
	    short lastCellNum = sheet.getRow(0).getLastCellNum();
	    System.out.println(lastCellNum);
	    String[][] data=new String[lastRowNum][lastCellNum];
	    
	    for (int i = 1; i <= lastRowNum; i++) {
	        for (int j = 0; j < lastCellNum; j++) {
	            String stringCellValue = sheet.getRow(i).getCell(j).getStringCellValue();
	       // System.out.println(stringCellValue);
	        data[i-1][j]=stringCellValue;
	        }
	    }
	    book.close();
	    return data;
	}
	}
