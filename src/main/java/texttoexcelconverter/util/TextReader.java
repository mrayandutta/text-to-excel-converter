package texttoexcelconverter.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import texttoexcelconverter.model.LineSplitInfo;

public class TextReader {
	
	public static void main(String args[])  
	{  
	try  
	{  
	//the file to be opened for reading  
	FileInputStream fis=new FileInputStream("C://data-validator-test//test.txt");       
	Scanner sc=new Scanner(fis);    //file to be scanned  
	//returns true if there is another line to read  
	while(sc.hasNextLine())  
	{  
	System.out.println(sc.nextLine());      //returns the line that was skipped  
	}  
	sc.close();     //closes the scanner  
	}  
	catch(IOException e)  
	{  
	e.printStackTrace();  
	}  
	}  
	
	
	public static List<List<String>> getTextDataAsList(String filePath,List<LineSplitInfo> splitInfoList) 
	{
		List<List<String>> allRowDataList = new ArrayList<>();
		
		try  
		{  
		//the file to be opened for reading  
		FileInputStream fis=new FileInputStream(filePath);       
		Scanner sc=new Scanner(fis);    //file to be scanned  
		//returns true if there is another line to read  
		while(sc.hasNextLine())  
		{  
			String line = sc.nextLine();
			List<String> rowDataList = new ArrayList<String>();
			//System.out.println("line:"+line);
			for(LineSplitInfo splitInfo:splitInfoList)
			{
				String data = line.substring(splitInfo.getStart(),splitInfo.getEnd());
				//System.out.println(data);
				rowDataList.add(data);
			}
			allRowDataList.add(rowDataList);
		//System.out.println(sc.nextLine());      //returns the line that was skipped  
		}  
		sc.close();     //closes the scanner  
		}  
		catch(IOException e)  
		{  
		e.printStackTrace();  
		}  
		
		return allRowDataList;
		
	}

}
