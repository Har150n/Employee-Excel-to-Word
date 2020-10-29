import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class EmployeeAL {
	private ArrayList<Employee> employeeAL;
	private FileInputStream fis;
	private static String picturePath = "C:\\Users\\Msi\\eclipse-workspace\\Employee Report\\Picture";
	private Workbook wb;
	private Sheet sh;

	//C:\Users\Msi\Downloads\Employee Excel to Word\Picture 
	private static void tableText(List<XWPFParagraph> paragraph, String name, String section, String ID, String[] fileNames) throws InvalidFormatException, IOException {
	    XWPFRun run = paragraph.get(0).createRun();
	    paragraph.get(0).setAlignment(ParagraphAlignment.CENTER);
	    boolean picExists = false;
	    for(int i = 0; i < fileNames.length; i++) {
	    	String fullID = ID + ".jpg";	
			if(fullID.equals(fileNames[i])) {
				InputStream pic = new FileInputStream(picturePath + "\\" + fileNames[i]);
				picExists = true;
				run.addPicture(pic, XWPFDocument.PICTURE_TYPE_JPEG, fileNames[i],Units.toEMU(60), Units.toEMU(72)); 			
			}
		}
	    if(!picExists) {
			InputStream pic = new FileInputStream("C:\\Users\\Msi\\eclipse-workspace\\Employee Report\\Picture\\Default Picture.png");
			run.addPicture(pic, XWPFDocument.PICTURE_TYPE_PNG, "Default Picture.png", Units.toEMU(60), Units.toEMU(72));
		}
			
		
	    run.addBreak();
	    run.setFontSize(11);
	    run.setFontFamily("Times New Roman");
	    run.setText(name);
	    run.addBreak();
	    run.setText(section);
	}
	
	public void extractEmployeeData(String location) throws Exception { //method to take the employee data from excel
		
		employeeAL = new ArrayList<>();
		String name;
		String department;
		String section;
		String ID;
		fis = new FileInputStream(location + ".xlsx");
		wb = WorkbookFactory.create(fis);
		sh = wb.getSheetAt(0); 
		int nameCol = 1;
		int departmentCol = 4;
		int sectionCol = 5;
		int IDCol = 2;
		for(Row row: sh) {
			if(row.getRowNum() == 0) {
				continue;
			}
			name = row.getCell(nameCol,MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
			department = row.getCell(departmentCol, MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
		    section = row.getCell(sectionCol, MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
		    ID = row.getCell(IDCol, MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue();
			employeeAL.add(new Employee(name, department, section, ID));
		    
		}
		
	}
	
	//returns a list of all the department categories
	public ArrayList<String> getDepartment(ArrayList<Employee> List) {
		ArrayList<String> departmentList = new ArrayList<>();
		
		for(int i = 0; i < List.size(); i++) {
			departmentList.add(List.get(i).getDepartment());
		}
		departmentList = (ArrayList<String>) departmentList.stream().distinct().collect(Collectors.toList()); //removes Dupes
		return departmentList;
	}
	
	public static String[] fileNames(String directoryPath) {

	    File dir = new File(directoryPath);

	    Collection<String> files  =new ArrayList<String>();

	    if(dir.isDirectory()){
	        File[] listFiles = dir.listFiles();

	        for(File file : listFiles){
	            if(file.isFile()) {
	                files.add(file.getName());
	            }
	        }
	    }

	    return files.toArray(new String[]{});
	}
	
	
	//writes the document 
	public void writeReport(ArrayList<Employee> List, String locationDoc, String locationPic) throws Exception {
		String[] files = fileNames(locationPic); //arrays of names of pictures in picture folder
		XWPFDocument doc = new XWPFDocument();
		FileOutputStream out = new FileOutputStream(locationDoc);
		int row = 0; 
		int column = 0;
		boolean firstEmployee = true;
		boolean checkFirst = true;
		ArrayList<String> departmentList = getDepartment(List);
		System.out.println("Running... ");
		for(String department : departmentList) {
			firstEmployee = true;
			checkFirst = true;
			row = 0;
			column = 0;
			XWPFParagraph docParagraph = doc.createParagraph();
			XWPFRun docRun = docParagraph.createRun();
			docParagraph.setAlignment(ParagraphAlignment.CENTER);
			docRun.setFontSize(28);
			docRun.setFontFamily("Times New Roman");
			docRun.setText(department );
			
			
			XWPFTable table = doc.createTable();
			table.removeBorders();
    		table.setTableAlignment(TableRowAlign.CENTER);
    		for(int i = 0; i < List.size(); i++) {
    			
    			String employeeDepartment = List.get(i).getDepartment();
    			String name = List.get(i).getName();
    			String section = List.get(i).getSection();
    			String ID = List.get(i).getID();
    			if(employeeDepartment.equals(department) ) {
    				if(firstEmployee && checkFirst) {
    					XWPFTableRow firstRow = table.getRow(0);
    					XWPFTableCell firstCell = firstRow.getCell(0);
    					tableText(firstCell.getParagraphs(), name, section, ID, files);
    					//firstCell.setText(info);
    					firstEmployee = false;    					
    					column++;
    				} else if(column <= 4 && checkFirst) {
    					XWPFTableRow firstRow = table.getRow(0);
    					//firstRow.addNewTableCell().setText(info);
    					tableText(firstRow.addNewTableCell().getParagraphs(), name, section, ID, files);
    					column++; 
    					
    				} else if(column >= 4 ){
    					column = 0;   					
    					checkFirst = false;
    					XWPFTableRow nextRow = table.createRow();
    					nextRow.setCantSplitRow(true);
    					
    					//creates new row
    					row++;
    					//nextRow.getCell(0).setText(info);
    					tableText(nextRow.getCell(0).getParagraphs(), name, section, ID, files);
    				} else {
    					//table.getRow(row).getCell(column).setText(info);
    					column++;
    					if(table.getRow(row).getCell(column) == null) 
    					{
    						table.getRow(row).createCell();
    					}
    					tableText(table.getRow(row).getCell(column).getParagraphs(), name, section, ID, files);					
    				
    			
    					

    				}
    			}
    		}
    		 doc.createParagraph().setPageBreak(true);
    	
			
		}
	
	        doc.write(out);
	        doc.close();
	        System.out.println("Finished Writing");
	}
	
	public ArrayList<Employee> getEmployeeAL() {
		return employeeAL;
	}
		
}
