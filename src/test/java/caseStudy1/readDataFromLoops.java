package caseStudy1;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class readDataFromLoops {

		@Test
		public void testdata() throws Exception {
			
			File src =new File("C:\\Users\\91956\\OneDrive\\Documents\\DXC\\eclipse\\secondMavenProject\\TestDataOrangeHrm\\TestDataOrg.xlsx");
			FileInputStream fis = new FileInputStream(src);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet1 = wb.getSheetAt(0);
			
			int rowcount = sheet1.getLastRowNum();
			System.out.println("total rows from excel sheet 1 is..."+ rowcount);
			
			for(int i=0; i<=rowcount; i++) {
				XSSFRichTextString data1 = sheet1.getRow(i).getCell(0).getRichStringCellValue();
				System.out.println("data from row 1 and index "+ i + " is " + data1);
				
				XSSFRichTextString data2 = sheet1.getRow(i).getCell(1).getRichStringCellValue();
				System.out.println("data from row 1 and index "+ i + " is " + data2);
				
				
			}
			
			wb.close();
			
		}
}
