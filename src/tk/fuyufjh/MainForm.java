package tk.fuyufjh;

import java.io.File;

import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.List;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;
import org.eclipse.wb.swt.SWTResourceManager;
import org.eclipse.swt.widgets.Label;

public class MainForm {
	private static Text textOpen;
	private static Text textSave;

	/**
	 * Launch the application.
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		Display display = Display.getDefault();
		final Shell shlCounselorAssistantBy = new Shell(SWT.MIN | SWT.CLOSE);
		shlCounselorAssistantBy.setSize(800, 600);
		shlCounselorAssistantBy.setText("Counselor Assistant");
		
		Group group1 = new Group(shlCounselorAssistantBy, SWT.NONE);
		group1.setText("\u7B2C\u4E00\u6B65\uFF1A\u9009\u62E9\u6587\u4EF6");
		group1.setBounds(10, 10, 774, 63);
		
		textOpen = new Text(group1, SWT.BORDER | SWT.READ_ONLY);
		textOpen.setBackground(SWTResourceManager.getColor(SWT.COLOR_WHITE));
		textOpen.setBounds(10, 25, 668, 23);
		
		Button btnOpenFile = new Button(group1, SWT.NONE);

		btnOpenFile.setBounds(684, 23, 80, 27);
		btnOpenFile.setText("\u6253\u5F00...");
		
		Group group2 = new Group(shlCounselorAssistantBy, SWT.NONE);
		group2.setText("\u7B2C\u4E8C\u6B65\uFF1A\u9009\u62E9\u8BFE\u7A0B");
		group2.setBounds(10, 89, 774, 329);
		
		final List listUnselected = new List(group2, SWT.BORDER | SWT.V_SCROLL);
		listUnselected.setBounds(10, 25, 334, 292);
		
		final List listSelected = new List(group2, SWT.BORDER | SWT.V_SCROLL);
		listSelected.setBounds(429, 25, 335, 292);
		
		Button btnSelect = new Button(group2, SWT.NONE);
		btnSelect.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				int index = listUnselected.getSelectionIndex();
				if (index < 0) return;
				Course c = Function.getUnselectedCourses().get(index);
				Function.getSelectedCourses().add(0, c);
				Function.getUnselectedCourses().remove(index);
				listSelected.add(c.toString(), 0);
				listUnselected.remove(index);
				if (index < Function.getUnselectedCourses().size())
					listUnselected.select(index);
			}
		});
		btnSelect.setBounds(350, 138, 73, 30);
		btnSelect.setText("---\u9009\u62E9-->");
		
		Button btnUnselect = new Button(group2, SWT.NONE);
		btnUnselect.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				int index = listSelected.getSelectionIndex();
				if (index < 0) return;
				Course c = Function.getSelectedCourses().get(index);
				Function.getUnselectedCourses().add(0, c);
				Function.getSelectedCourses().remove(index);
				listUnselected.add(c.toString(), 0);
				listSelected.remove(index);
				if (index < Function.getSelectedCourses().size())
					listSelected.select(index);
			}
		});
		btnUnselect.setText("<--\u6392\u9664---");
		btnUnselect.setBounds(350, 188, 73, 30);
		
		Group group3 = new Group(shlCounselorAssistantBy, SWT.NONE);
		group3.setText("\u7B2C\u4E09\u6B65\uFF1A\u5BFC\u51FA\u6587\u4EF6");
		group3.setBounds(10, 432, 774, 107);
		
		textSave = new Text(group3, SWT.BORDER | SWT.READ_ONLY);
		textSave.setBackground(SWTResourceManager.getColor(SWT.COLOR_WHITE));
		textSave.setBounds(10, 29, 668, 23);
		
		Button buttonSelectSavePath = new Button(group3, SWT.NONE);
		buttonSelectSavePath.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				FileDialog fd = new FileDialog(shlCounselorAssistantBy,SWT.SAVE);
				fd.setFilterPath(System.getProperty("JAVA.HOME"));
				fd.setFilterExtensions(new String[]{"*.xls"});
				fd.setFilterNames(new String[]{"Microsoft Excel 文档"});
				fd.setFileName(textSave.getText());
				String file = fd.open();
				if (file != null)
					textSave.setText(file);
			}
		});
		buttonSelectSavePath.setText("\u9009\u62E9...");
		buttonSelectSavePath.setBounds(684, 27, 80, 27);
		
		Button btnExport = new Button(group3, SWT.NONE);
		btnExport.setFont(SWTResourceManager.getFont("华文细黑", 11, SWT.NORMAL));
		btnExport.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				try {
					Function.exportExcel(textSave.getText());
					
					MessageBox messageBox = new MessageBox(shlCounselorAssistantBy, SWT.OK);
					messageBox.setText("Done");
					messageBox.setMessage("保存成功！");
					messageBox.open();
				} catch (Exception ex) {
					MessageBox messageBox = new MessageBox(shlCounselorAssistantBy, SWT.OK);
					messageBox.setText("Error");
					messageBox.setMessage("文件保存失败！请检查文件是否在别的地方打开？或尝试更换文件名。\n\n"+ex.getMessage());
					messageBox.open();
					ex.printStackTrace();
				}
			}
		});
		btnExport.setBounds(303, 67, 163, 30);
		btnExport.setText("\u5BFC\u51FA");
		
		Label lblFuyufjh = new Label(shlCounselorAssistantBy, SWT.NONE);
		lblFuyufjh.setEnabled(false);
		lblFuyufjh.setForeground(SWTResourceManager.getColor(SWT.COLOR_WHITE));
		lblFuyufjh.setAlignment(SWT.RIGHT);
		lblFuyufjh.setBounds(426, 545, 358, 17);
		lblFuyufjh.setText("Klose@NJU  fuyufjh@163.com");
		
		btnOpenFile.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				FileDialog fd = new FileDialog(shlCounselorAssistantBy,SWT.OPEN);
				fd.setFilterPath(System.getProperty("JAVA.HOME"));
				fd.setFilterExtensions(new String[]{"*.xls"});
				fd.setFilterNames(new String[]{"Microsoft Excel 文档"});
				String file = fd.open();
				if (file != null) {
					textOpen.setText(file);
					textSave.setText(file.replace(".xls", "_整理.xls"));
					textOpen.setBackground(SWTResourceManager.getColor(160, 250, 150));
					
					try {
						Function.fileOpen = new File(file);
						Function.prepareDatebase();
						Function.ReadExcel();
						
						Function.prepareCourseList();
						for (Course c: Function.getSelectedCourses()) {
							listSelected.add(c.toString());
						}
						for (Course c: Function.getUnselectedCourses()) {
							listUnselected.add(c.toString());
						}
					}
					catch (Exception ex) {
						MessageBox messageBox = new MessageBox(shlCounselorAssistantBy, SWT.OK);
						messageBox.setText("Error");
						messageBox.setMessage("文件加载失败！\n\n请检查文件格式以及是否可读。\n请检查是否打开了多个本程序。\n\n"+ex.getMessage());
						messageBox.open();
						ex.printStackTrace();
					}
				}
			}
		});

		shlCounselorAssistantBy.open();
		shlCounselorAssistantBy.layout();
		while (!shlCounselorAssistantBy.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
	}
}
