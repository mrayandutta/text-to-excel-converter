package texttoexcelconverter.util;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;

import texttoexcelconverter.model.LineSplitInfo;

public class StringUtil {
	
	public static void main(String[] args) {
		String line = "00001111000009ABC";
		List<String> columnLengths  = Arrays.asList("4", "4","6","3");
		int start = 0;
		int end = 0;
		for (Iterator iterator = columnLengths.iterator(); iterator.hasNext();) {
			String lengthStr = (String) iterator.next();
			end = end + Integer.valueOf(lengthStr) ;
			LineSplitInfo splitInfo = new LineSplitInfo(start, end);
			System.out.println(line.substring(start,end));
			start = end;
			System.out.println(splitInfo);
			
			
		}
		
	}
	
	public static List<LineSplitInfo> getValidationMappingData(List<String> columnLengths)
    {
		List<LineSplitInfo> lineSplitInfoList = new ArrayList<LineSplitInfo>();
		int start = 0;
		int end = 0;
		for (Iterator iterator = columnLengths.iterator(); iterator.hasNext();) {
			String lengthStr = (String) iterator.next();
			end = end + Integer.valueOf(lengthStr) ;
			LineSplitInfo splitInfo = new LineSplitInfo(start, end);
			lineSplitInfoList.add(splitInfo);
			start = end;
			
		}
		return lineSplitInfoList;
    }

}
