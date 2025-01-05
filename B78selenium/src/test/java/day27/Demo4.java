package day27;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo4 {

	public static void main(String[] args) throws Exception {

		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/book1.xlsx"));
		String sheet="Sheet1";
		wb.getSheet(sheet).getRow(0).getCell(0).setCellValue("Bhanu");
		wb.write(new FileOutputStream("./data/book2.xlsx"));
		wb.close();
		System.out.println("Done");
	}

}
