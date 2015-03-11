package tk.fuyufjh;

public class Student {
	private String code;
	private String name;
	private double gpa;
	
	public Student()
	{
		
	}
	
	public Student(String code, String name)
	{
		this.code = code;
		this.name = name;
	}
	
	public String getCode()
	{
		return code;
	}
	
	public String getName()
	{
		return name;
	}
	
	public void setGPA(double gpa)
	{
		this.gpa = gpa;
	}
	
	public double getGPA()
	{
		return gpa;
	}
}
