package Excel.Sec26_Excel;

import java.io.IOException;
import java.util.ArrayList;

public class dataDriven{
	
	public static void main(String[] args) throws IOException {
		
		//Lec171_AccessExcel d=new Lec171_AccessExcel();
		AccessExcel d=new AccessExcel();
		ArrayList a=d.getDate("Delete profile");  //the method will return arraylist as data
		
		System.out.println(a.get(0));
		System.out.println(a.get(1));
		System.out.println(a.get(2));
		System.out.println(a.get(3));
		
	}
}