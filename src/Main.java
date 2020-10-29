import java.util.ArrayList;
import java.util.Scanner;


public class Main {
	
	public static void mainMethod(String locationEx, String locationDoc, String locationPic) throws Exception {
		EmployeeAL arrObj = new EmployeeAL();
		arrObj.extractEmployeeData(locationEx);
		
		ArrayList<Employee> List = arrObj.getEmployeeAL();
		//C:\Users\Msi\Downloads\Employee Excel to Word\HR Shwe Palin CV Data
		arrObj.writeReport(List, locationDoc + "\\Output.docx", locationPic );
		
	}
	

}
