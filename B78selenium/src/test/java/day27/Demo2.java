package day27;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo2 {

	public static void main(String[] args) throws Exception {

		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/book1.xlsx"));
		int rc=wb.getSheet("sheet1").getLastRowNum();//last index of the row
		for(int i=0;i<=rc;i++)
		{
			int cc=wb.getSheet("sheet1").getRow(i).getLastCellNum();//column count
			for(int j=0;j<cc;j++) {
				String v=wb.getSheet("sheet1").getRow(i).getCell(j).toString();
				System.out.print(v+" ");
			}
			System.out.println();
		}
		wb.close();
	}

}
