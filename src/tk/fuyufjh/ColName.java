package tk.fuyufjh;

public class ColName {

	/**
	 * ��ȡһ�ж�Ӧ����ĸ�����磺ColumnNum=1���򷵻�ֵΪA �к�ת��ĸ
	 */
	public static String getColumnName(int columnNum) {
		int first;
		int last;
		String result = "";
		if (columnNum > 256)
			columnNum = 256;
		first = columnNum / 27;
		last = columnNum - (first * 26);

		if (first > 0)
			result = String.valueOf((char) (first + 64));

		if (last > 0)
			result = result + String.valueOf((char) (last + 64));

		return result;
	}

	public static int nameToColumn(String name) {
		int column = -1;
		for (int i = 0; i < name.length(); ++i) {
			int c = name.charAt(i);
			column = (column + 1) * 26 + c - 'A';
		}
		return column;
	}
}
