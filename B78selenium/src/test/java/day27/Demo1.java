package day27;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo1 {

	public static void main(String[] args) throws EncryptedDocumentException, FileNotFoundException, IOException {
		//open the excel file
		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/book1.xlsx"));
		
		for(int i=0;i<4;i++)
		{
			for(int j=0;j<3;j++) {
				String v=wb.getSheet("sheet1").getRow(i).getCell(j).toString();
				System.out.print(v+" ");
			}
			System.out.println();
		}
		wb.close();
	}

}
