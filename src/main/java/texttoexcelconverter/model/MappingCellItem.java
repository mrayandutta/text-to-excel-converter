package texttoexcelconverter.model;

public class MappingCellItem {
	
	public MappingCellItem(String columnName, int columnLength) {
		super();
		this.columnName = columnName;
		this.columnLength = columnLength;
	}
	
	private String columnName;
	private int columnLength;
	public String getColumnName() {
		return columnName;
	}
	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}
	public int getColumnLength() {
		return columnLength;
	}
	public void setColumnLength(int columnLength) {
		this.columnLength = columnLength;
	}
	
	@Override
	public String toString() {
		return "MappingCellItem [columnName=" + columnName + ", columnLength=" + columnLength + "]";
	}

	
	
	

}
