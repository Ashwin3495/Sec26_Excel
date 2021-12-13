package Excel.Sec26_Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AccessExcel {

	public ArrayList <String> getDate(String testcasename) throws IOException{
		// TODO Auto-generated method stub
				ArrayList <String>a=new ArrayList<String>();
				FileInputStream fis=new FileInputStream("C://Tools//Excel.xlsx");
				XSSFWorkbook wb=new XSSFWorkbook(fis);
				
				String sheets=wb.getSheetName(0);
				int sheet=wb.getNumberOfSheets();
				System.out.println(sheet);
				
				for(int i=0;i<sheet;i++)
				{
					if(wb.getSheetName(i).equalsIgnoreCase("Demo"))
					{
						
						XSSFSheet sheet1=wb.getSheetAt(i); //Access the sheet first cell and wait
						Iterator<Row> rows=sheet1.iterator(); //will move to next row
						Row firstrow=rows.next();
						Iterator<Cell> ce=firstrow.cellIterator();
						int column=0;
						int k=0;
						while(rows.hasNext())
						{
							Row value=rows.next();
							if(value.getCell(k).getStringCellValue().equalsIgnoreCase("Delete profile"))
							{
								column=k;
							}
							k++;
						}
						System.out.println(k);
						
						Row r=rows.next();
						Iterator<Cell> cv = null;
								while(cv.hasNext())
								{
									Cell c=cv.next();
									if(c.getCellType()==CellType.STRING)
									{
										a.add(c.getStringCellValue());
									}
									else {
										a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
									}
									a.add(cv.next().getStringCellValue());
								}
							}
							
						}
					
				return a;
	}
	public static void main(String[] args) throws IOException {
		
		
	}

}
