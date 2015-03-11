package tk.fuyufjh;

public class Course {
	private String code;
	private String name;
	private int count;
	private int score;
	private String semester;
	
	public Course() {
		// TODO Auto-generated constructor stub
	}
	
	public Course(String code, String name, int count, int score, String semester) {
		this.code = code;
		this.name = name;
		this.count = count;
		this.score = score;
		this.semester = semester;
	}
	
	public String getCode()
	{
		return this.code;
	}
	
	public String getName()
	{
		return this.name;
	}
	
	public String getSemester()
	{
		return this.semester;
	}
	
	public int getCount()
	{
		return this.count;
	}
	
	public int getScore()
	{
		return this.score;
	}
	
	public void setCount(int n)
	{
		count = n;
	}

	@Override
	public String toString() {
		return "[" + code + "] " + name;
	}
}
