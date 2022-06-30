package texttoexcelconverter.model;

public class LineSplitInfo {
	private Integer start ;
	private Integer end ;
	
	
	
	
	public LineSplitInfo(Integer start, Integer end) {
		super();
		this.start = start;
		this.end = end;
	}
	public Integer getStart() {
		return start;
	}
	public void setStart(Integer start) {
		this.start = start;
	}
	public Integer getEnd() {
		return end;
	}
	public void setEnd(Integer end) {
		this.end = end;
	}
	@Override
	public String toString() {
		return "LineSplitInfo [start=" + start + ", end=" + end + "]";
	}
	
	

}
