package texttoexcelconverter.model;

import org.apache.poi.ss.usermodel.Cell;

public class CellItem
{
    private String data;
    private int rowNumber;
    private int columnNumber;
    private Cell cell;

    public CellItem(String data, int rowNumber, Cell cell) {
        this.data = data;
        this.rowNumber = rowNumber;
        this.cell = cell;
    }
    
    

    public CellItem(String data, int rowNumber, int columnNumber, Cell cell) {
		super();
		this.data = data;
		this.rowNumber = rowNumber;
		this.columnNumber = columnNumber;
		this.cell = cell;
	}


	public int getColumnNumber() {
		return columnNumber;
	}



	public void setColumnNumber(int columnNumber) {
		this.columnNumber = columnNumber;
	}



	public String getData() {
        return data;
    }

    public void setData(String data) {
        this.data = data;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public void setRowNumber(int rowNumber) {
        this.rowNumber = rowNumber;
    }

    public Cell getCell() {
        return cell;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }

    @Override
    public String toString() {
        return "CellItem{" +
                "data='" + data + '\'' +
                ", rowNumber=" + rowNumber +
                ", cell=" + cell +
                '}';
    }
}
