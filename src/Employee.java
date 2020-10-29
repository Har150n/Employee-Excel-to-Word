

public class Employee {
	
	private String name;
	private String department;
	private String section;
	private String ID;

	public Employee(String name, String department, String job, String ID) { //constructor for employee
		this.name = name;
		this.department = department;
		section = job;
		this.ID = ID;
	}
	
	public String getName() {
		return name;
	}
	public String getDepartment() {
		return department;
	}
	public String getSection() {
		return section;
	}
	public String getID() {
		return ID;
	}
}


