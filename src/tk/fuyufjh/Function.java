package tk.fuyufjh;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Function {

	static File fileOpen;
	static Connection sqlConnection;
	static Statement sqlStatement;
	static private ArrayList<Student> student;
	static private ArrayList<Course> listSelectedCourses;
	static private ArrayList<Course> listUnselectedCourses;
	
	static void prepareDatebase() throws Exception
	{
        Class.forName("org.sqlite.JDBC");
        sqlConnection = DriverManager.getConnection("jdbc:sqlite:db");
        sqlStatement = sqlConnection.createStatement();
        sqlStatement.executeUpdate("drop table if exists jw;");
	}
	
	static void ReadExcel() throws Exception
	{
		Workbook workbook = Workbook.getWorkbook(fileOpen);
		Sheet sheet = workbook.getSheet(0);
		
		// First line is table head.
		int nCol = sheet.getColumns();
		int nRow = sheet.getRows();
		String columnName = null;
		StringBuilder sql = new StringBuilder("create table 'jw' (");
		boolean first = true;
		for (int i=0; i<nCol; i++) {
			columnName = sheet.getCell(i, 0).getContents();
			sql.append((first?"'":",'")  + columnName + "'");
			first = false;
		}
		sql.append(");");
		//System.out.println(sql.toString());
		sqlStatement.executeUpdate(sql.toString());
				
		// insert into jw values('ZhangSan',8000);

		for (int i=1; i<nRow; i++) {
			String studentNo = sheet.getCell(0, i).getContents();
			sql = new StringBuilder("insert into jw values(");
			sql.append("'"+ studentNo +"'");
			for (int j=1; j<nCol; j++) {
				sql.append(",'"+ sheet.getCell(j, i).getContents() +"'");
			}
			sql.append(");");
			sqlStatement.executeUpdate(sql.toString());
			//System.out.println(sql.toString());
		}
		//sqlConnection.close();
		workbook.close();
	}
	
	static void prepareCourseList()
	{
		listSelectedCourses = new ArrayList<Course>();
		listUnselectedCourses = new ArrayList<Course>();
		student = new ArrayList<Student>();
		try {
			// Add students
			ResultSet rs = sqlStatement.executeQuery("select distinct 学号,姓名 from jw;");
			while (rs.next()) {
				student.add(new Student(rs.getString(1), rs.getString(2)));
			}
			
			// Add courses
			rs = sqlStatement.executeQuery("select 课程编号,课程名称,学分,学期 from jw group by 课程编号,学期");
			String prevNoCourse = "";
			while (rs.next()) {
				if (prevNoCourse.equals(rs.getString(1))) {
					continue; //重修的。。。尼玛
				}
				listUnselectedCourses.add(new Course(rs.getString(1), rs.getString(2), 0 , rs.getInt(3), rs.getString(4)));
				prevNoCourse = rs.getString(1);
			}
			for (Course c: listUnselectedCourses) {
				rs = sqlStatement.executeQuery("select count(case 课程编号 when '"
					 + c.getCode() + "' then 1 end) from jw;");
				rs.next();
				int count = rs.getInt(1);
				c.setCount(count);
			}
			
			for (int i=0; i<listUnselectedCourses.size();) {
				if (listUnselectedCourses.get(i).getCount() == student.size()) {
					listSelectedCourses.add(listUnselectedCourses.get(i));
					listUnselectedCourses.remove(i);
				}
				else i++;
			}
			
			Collections.sort(listUnselectedCourses, new Comparator<Course>() {

				@Override
				public int compare(Course o1, Course o2) {
					return o2.getCount() - o1.getCount();
				}
				
			});
			
			rs.close();
		}
		catch (Exception ex) {
			ex.printStackTrace();
		}

		
	}
	
	static ArrayList<Course> getSelectedCourses()
	{
		return listSelectedCourses;
	}
	
	static ArrayList<Course> getUnselectedCourses()
	{
		return listUnselectedCourses;
	}
	
	static void exportExcel(String f) throws Exception
	{
		ArrayList<Course> copyCourseList = new ArrayList<Course>(listSelectedCourses);
		ArrayList<Student> copyStudentList = new ArrayList<Student>(student);
		Collections.sort(copyCourseList, new Comparator<Course>() {

			@Override
			public int compare(Course o1, Course o2) {
				return o1.getSemester().compareTo(o2.getSemester());
			}
		});
		
		WritableWorkbook workbook = Workbook.createWorkbook(new File(f));
		WritableSheet sheet = workbook.createSheet("成绩统计", 0);
		WritableCellFormat wcFormat = new WritableCellFormat();
		wcFormat.setAlignment(Alignment.CENTRE);
		wcFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		sheet.addCell(new Label(0, 0, "学号", wcFormat));
		sheet.setColumnView(0, 11);
		sheet.addCell(new Label(1, 0, "姓名", wcFormat));
		int i=2;
		wcFormat = new WritableCellFormat();
		wcFormat.setAlignment(Alignment.CENTRE);
		wcFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		wcFormat.setWrap(true);
		
		WritableCellFormat wcFormatCenterC1 = new WritableCellFormat();
		wcFormatCenterC1.setVerticalAlignment(VerticalAlignment.CENTRE);
		wcFormatCenterC1.setAlignment(Alignment.CENTRE);
		wcFormatCenterC1.setBackground(Colour.IVORY);
		wcFormatCenterC1.setWrap(true);
		
		WritableCellFormat wcFormatCenterC2 = new WritableCellFormat();
		wcFormatCenterC2.setVerticalAlignment(VerticalAlignment.CENTRE);
		wcFormatCenterC2.setAlignment(Alignment.CENTRE);
		wcFormatCenterC2.setBackground(Colour.LIGHT_TURQUOISE);
		wcFormatCenterC2.setWrap(true);
		
		String prev_sem = "";
		for (Course c:copyCourseList) {
			if (!prev_sem.equals(c.getSemester())) {
				WritableCellFormat t = wcFormatCenterC2;
				wcFormatCenterC2 = wcFormatCenterC1;
				wcFormatCenterC1 = t;
			}
			sheet.addCell(new Label(i++, 0, c.getName() + " " + c.getScore() + "学分", wcFormatCenterC1));
			prev_sem = c.getSemester();
		}
		sheet.addCell(new Label(i++, 0, "学分数", wcFormat));
		sheet.addCell(new Label(i++, 0, "学分绩", wcFormat));
		sheet.addCell(new Label(i++, 0, "专业排名", wcFormat));
		
		//=====================================
		int line = 1;
		WritableCellFormat wcFormatLeft = new WritableCellFormat();
		wcFormatLeft.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		WritableCellFormat wcFormatCenter = new WritableCellFormat();
		wcFormatCenter.setVerticalAlignment(VerticalAlignment.CENTRE);
		wcFormatCenter.setAlignment(Alignment.CENTRE);
		
		NumberFormat nf2d = new jxl.write.NumberFormat("#.00");
		WritableFont fontRegular = new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD,
				false, UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
		WritableCellFormat wcFormatCenter2D = new WritableCellFormat(fontRegular, nf2d);
		wcFormatCenter2D.setVerticalAlignment(VerticalAlignment.CENTRE);
		wcFormatCenter2D.setAlignment(Alignment.CENTRE);
		
		WritableCellFormat wcFormatCenterBJG = new WritableCellFormat();
		wcFormatCenterBJG.setVerticalAlignment(VerticalAlignment.CENTRE);
		wcFormatCenterBJG.setAlignment(Alignment.CENTRE);
		WritableFont fontRed = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD,
				false, UnderlineStyle.NO_UNDERLINE, Colour.RED);
		wcFormatCenterBJG.setFont(fontRed);
		
		for (Student s: copyStudentList) {
			sheet.addCell(new Label(0, line, s.getCode(), wcFormatLeft));
			sheet.addCell(new Label(1, line, s.getName(), wcFormatLeft));
			i = 2;
			int sumCredit = 0;
			int total = 0;
			StringBuilder form = new StringBuilder("(");
			boolean first=true;
			for (Course c:copyCourseList) {
				ResultSet rs = sqlStatement.executeQuery("select 总评,学分 from jw where 学号='"
						       +s.getCode()+"' AND 课程编号='"+c.getCode()+"' order by 学期 DESC;");
				// 以最后一次成绩计算
				int score = 0;
				if (rs.next()) score=Math.abs(rs.getInt(1)); 
				if (score != 0) {
					if (score >= 60)
						sheet.addCell(new Label(i++, line, Integer.toString(score), wcFormatCenter));
					else 
						sheet.addCell(new Label(i++, line, Integer.toString(score), wcFormatCenterBJG));
					sumCredit += rs.getInt(2);
					form.append((first?"":"+")+ColName.getColumnName(i) + (line+1) +"*"+rs.getInt(2));
					total += rs.getInt(2) * rs.getInt(1);
					first = false;
				}
				else {
					sheet.addCell(new Label(i++, line, ""));
				}
			}
			sheet.addCell(new Label(i++, line, Integer.toString(sumCredit), wcFormatCenter));
			form.append(")/" + ColName.getColumnName(i) + (line+1));
			sheet.addCell(new Formula(i++, line, form.toString(), wcFormatCenter2D));
			sheet.setRowView(line, 500);
			line++;
			s.setGPA((double) total / sumCredit);
		}
		
		Collections.sort(copyStudentList, new Comparator<Student>() {
			@Override
			public int compare(Student o1, Student o2) {
				return (int) ((o2.getGPA() - o1.getGPA())*1000.0);
			}
		});
		int colRank = 4 + copyCourseList.size();
		for (int k=0; k<student.size(); k++) {
			sheet.addCell(new Label(colRank, k+1, Integer.toString(copyStudentList.indexOf(student.get(k))+1), wcFormatCenter));
		}
		
		workbook.write();
		workbook.close();
	}
	
}
