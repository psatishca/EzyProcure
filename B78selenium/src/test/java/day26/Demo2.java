package day26;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo2 {

	public static void main(String[] args) throws Exception {
		String path="./data/Selenium.xlsx";
		
		//open the xl file
		FileInputStream fis=new FileInputStream(path);
		Workbook wb = WorkbookFactory.create(fis);
		//goto sheet1
		Sheet sheet = wb.getSheet("Sheet1");
		//goto 1sr row
		Row r = sheet.getRow(0);
		//goto 1st cell
		Cell c = r.getCell(0);
		//get cell value in String format & print it
		String v = c.getStringCellValue();
		System.out.println(v);
		//close the xl file
		wb.close();

	}

}
