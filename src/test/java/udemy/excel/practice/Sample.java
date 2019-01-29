package udemy.excel.practice;

import java.io.IOException;
import java.util.ArrayList;

public class Sample {
	
	public static void main(String args[]) throws IOException{
		
		ExcelDataDrive data = new ExcelDataDrive();
		ArrayList<String> getdata =data.getData("SignUp");
		System.out.println(getdata.get(0));
	}
}
