package day27;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo3 {

	public static void main(String[] args) throws Exception {

		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/book1.xlsx"));
		String sheet="Sheet2";
		int rc=wb.getSheet(sheet).getLastRowNum();//last index of the row
		for(int i=0;i<=rc;i++)
		{
		
			try 
			{
					int cc=wb.getSheet(sheet).getRow(i).getLastCellNum();//column count
		
						for(int j=0;j<cc;j++) 
						{
								try
								{
									String v=wb.getSheet(sheet).getRow(i).getCell(j).toString();
									System.out.print(v+" ");
								}
								catch (Exception e) {
									System.out.print("-- ");
								}
						}
			}
			catch (Exception e) 
			{
				System.out.print("-- -- --");
			}
			System.out.println();
		}
		wb.close();
	}

}
