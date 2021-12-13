package Excel.Sec26_Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Lec171_AccessExcel {
	
	public ArrayList <String> getData(String testcasename) throws IOException
	{

		ArrayList <String>a=new ArrayList<String>();
		FileInputStream fis=new FileInputStream("C://Tools//Excel.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		
		//Get the no.of sheet and match sheetname  --1
		int sheets=workbook.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{
			
			if(workbook.getSheetName(i).equalsIgnoreCase("Demo"))  //check if the sheet name is equal to Demo
			{
				XSSFSheet sheet=workbook.getSheetAt(i); //This will access the first cell 
				System.out.print(sheet.getFirstRowNum());
				
				Iterator<Row> rows=sheet.iterator(); //sheet is collection of rows , move to column
				Row firstrow=rows.next();
				Iterator<Cell> ce=firstrow.cellIterator(); //row is collection of cells, move to every cell
				int k=0;
				int coloumn = 0;
				
				//Identify testcase column by scanning entire 1st row
				while(ce.hasNext()) //if next cell is present, ce is used for cell sideways
				{
					Cell value=ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("Testcases"))
					{
						coloumn=k; //Store the Testcases index value
					}
					k++;
				}
				
				System.out.println(coloumn);
				
				//Once column is identified then scan entire testcase coloumn to identify purchase testcase row 
				
				while(rows.hasNext()) //move to next rows and check data is present, move downwards
				{
					Row r=rows.next(); //will move to next row to store value
					if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcasename))  //checks if given value is present
					{
						//after purchase is found grab all data of that row
						Iterator<Cell> cv=r.cellIterator();
						while(cv.hasNext())
						{
							Cell c=cv.next();
							if(c.getCellType()==CellType.STRING)
							{
								a.add(c.getStringCellValue());
							}
							else {
							a.add(NumberToTextConverter.toText(c.getNumericCellValue())); //convert numeric to string
								
							}
							//System.out.println(cv.next().getStringCellValue());
							a.add(cv.next().getStringCellValue());
						}
						
					}
				}
			}
		}
		//---1		
		return a;
		
		
	}

	public static void main(String[] args) throws IOException {
		
		
		
	}

}
