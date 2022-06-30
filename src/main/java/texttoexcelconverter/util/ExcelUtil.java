package texttoexcelconverter.util;


import java.io.File;
import java.io.FileOutputStream;
import java.lang.invoke.MethodHandles;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import texttoexcelconverter.model.CellItem;
import texttoexcelconverter.model.LineSplitInfo;
import texttoexcelconverter.model.MappingCellItem;

public class ExcelUtil {
    public static final Logger logger = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass());

    /**
     * Method to return an Excel workbook book when location is supplied
     * @param filePath
     * @return
     */
    public static Workbook getWorkbookFromExcel(String filePath)
    {
        Workbook workbook  = null;
        try {
            workbook = WorkbookFactory.create(new File(filePath));
        }
        catch (Exception e)
        {
        	logger.error("Error while trying to get workbook from filePath*:{}",filePath);
        	logger.error("Exception details :{}",e);
            //e.printStackTrace();
        }
        return workbook;
    }

    public static Sheet getSheetFromWorkbook(Workbook workbook,int sheetNumber)
    {
       return workbook.getSheetAt(sheetNumber);
    }
    
    
    public static Map<Integer,Map> getDataFromSheetCustomized(Sheet sheet,List<String> columnList)
    {   
    	Map<String,String> headerMap = new LinkedHashMap<>();
    	for (int j = 0; j < columnList.size(); j++)
        {
    		headerMap.put(String.valueOf(j), columnList.get(j));
        }
    	
    	Map<Integer,Map> rowMap = new LinkedHashMap<>();
        try
        {
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            // Create a DataFormatter to format and get each cell's value as String
            DataFormatter dataFormatter = new DataFormatter();
            for (int i = 0; i < numberOfRows; i++)
            {
                Map<String,CellItem> dataMap = new LinkedHashMap<>();
                Row row = sheet.getRow(i);
                int numberOfColumns = row.getPhysicalNumberOfCells();
                for (int j = 0; j < numberOfColumns; j++)
                {
                    Cell cell = row.getCell(j);
                    if(cell!=null)
                    {
                        String cellValue = dataFormatter.formatCellValue(cell);
                        String columnName= headerMap.get(String.valueOf(j));
                        CellItem cellItem = new CellItem(cellValue,i,cell);
                        dataMap.put(columnName,cellItem);
                    }
                    else
                    {
                        //logger.info("cell is null");
                    }
                }
                if(i!=0)
                {
                    rowMap.put(i,dataMap);
                }
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
           //
        }
        //logger.info("Read data from sheet '"+sheetName+"'");
        logger.info("rowMap:{}",rowMap);
        return rowMap;
    }

    public static Map<Integer,Map> getDataFromSheet(Sheet sheet)
    {   Map<Integer,Map> rowMap = new LinkedHashMap<>();
        try
        {
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            Map<String,String> headerMap = new LinkedHashMap<>();
            // Create a DataFormatter to format and get each cell's value as String
            DataFormatter dataFormatter = new DataFormatter();
            for (int i = 0; i < numberOfRows; i++)
            {
                Map<String,CellItem> dataMap = new LinkedHashMap<>();
                Row row = sheet.getRow(i);
                int numberOfColumns = row.getPhysicalNumberOfCells();
                for (int j = 0; j < numberOfColumns; j++)
                {
                    Cell cell = row.getCell(j);
                    if(cell!=null)
                    {
                        String cellValue = dataFormatter.formatCellValue(cell);
                        if(i==0)
                        {
                            headerMap.put(String.valueOf(j),cellValue);
                        }
                        else
                        {
                            String columnName= headerMap.get(String.valueOf(j));
                            CellItem cellItem = new CellItem(cellValue,i,cell);
                            dataMap.put(columnName,cellItem);
                        }
                    }
                    else
                    {
                        //logger.info("cell is null");
                    }
                }
                if(i!=0)
                {
                    rowMap.put(i,dataMap);
                }
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
           //
        }
        //logger.info("Read data from sheet '"+sheetName+"'");
        logger.info("rowMap:{}",rowMap);
        return rowMap;
    }
    
    public static List<String> getColumnListFromSheet(Sheet sheet)
    {   
    	return getRowDataListFromSheet(sheet, 0);

    }
    
    public static List<String> getRowDataListFromSheet(Sheet sheet,int rowNumber)
    {   List<String> rowDataList = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();
        Row row = sheet.getRow(rowNumber);
        int numberOfColumns = row.getPhysicalNumberOfCells();
        for (int j = 0; j < numberOfColumns; j++)
        {
            Cell cell = row.getCell(j);
            String cellValue = dataFormatter.formatCellValue(cell);
            rowDataList.add(cellValue);
        }
        
        
        logger.info("rowDataList:{}",rowDataList);
        return rowDataList;

    }
    
    public static List<MappingCellItem> getValidationMappingDataListFromSheet(Sheet sheet)
    {   
    	List<String> columnNameDataList = getRowDataListFromSheet(sheet, 0);
    	List<String> columnLengthDataList = getRowDataListFromSheet(sheet, 1);
    	List<MappingCellItem> validationDataList = new ArrayList<>();
        
        for (int j = 0; j < columnNameDataList.size(); j++)
        {
            MappingCellItem mappingCellItem = new MappingCellItem(columnNameDataList.get(j), Integer.valueOf(columnLengthDataList.get(j)));
            validationDataList.add(mappingCellItem);
        }
        
        
        logger.info("validationDataList:{}",validationDataList);
        return validationDataList;

    }

    public static Map<String,String> getMappingData(Sheet sheet)
    {
        Map<String,String> sourceTargetMap = new LinkedHashMap<>();
        try
        {
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            List<Integer> keyColumnIndexList = new ArrayList<>();
            Map<Integer,String> sourceColumnMap = new LinkedHashMap<>();
            Map<Integer,String> targetColumnMap = new LinkedHashMap<>();
            // Create a DataFormatter to format and get each cell's value as String
            DataFormatter dataFormatter = new DataFormatter();
            for (int i = 0; i < numberOfRows; i++)
            {
                Map<String,CellItem> dataMap = new LinkedHashMap<>();
                Row row = sheet.getRow(i);
                int numberOfColumns = row.getPhysicalNumberOfCells();
                for (int j = 0; j < numberOfColumns; j++)
                {
                    Cell cell = row.getCell(j);
                    if(cell!=null)
                    {
                        String cellValue = dataFormatter.formatCellValue(cell);
                        if(i==0)
                        {
                           if( j >0 && cellValue!=null && cellValue.equalsIgnoreCase("Y"))
                           {
                               keyColumnIndexList.add(j);
                           }
                        }
                        else
                        {
                            if(i==1 & j>0)
                            {
                                sourceColumnMap.put(j,cellValue);
                            }
                            if(i==2 & j>0)
                            {
                                targetColumnMap.put(j,cellValue);
                                sourceTargetMap.put(sourceColumnMap.get(j),cellValue);
                            }
                        }
                    }
                }
            }

        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            //
        }
        //logger.info("Read data from sheet '"+sheetName+"'");
        logger.info("sourceTargetMap:{}",sourceTargetMap);
        return sourceTargetMap;

    }

    public static void highLightCell(Cell cell,Workbook workbook) {
        // Create a Font for styling header cells
        //Font headerFont = workbook.createFont();
        //headerFont.setBold(true);
        //headerFont.setFontHeightInPoints((short) 14);
        //headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        //headerCellStyle.setFont(headerFont);

        // fill foreground color ...
        headerCellStyle.setFillForegroundColor(IndexedColors.YELLOW.index);
        // and solid fill pattern produces solid grey cell fill
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(headerCellStyle);

    }
    
    public static void addCommentToCell(Workbook workbook,int sheetNumber,Cell cell,String commentText) 
    {
    	Sheet sheet = workbook.getSheetAt(sheetNumber);
    	Drawing<Shape> drawing = (Drawing<Shape>) sheet.createDrawingPatriarch();
		ClientAnchor clientAnchor = drawing.createAnchor(0, 0, 0, 0, 0, 2, 7, 12);

		Comment comment = (Comment) drawing.createCellComment(clientAnchor);
		CreationHelper creationHelper = (XSSFCreationHelper) workbook.getCreationHelper();	
		RichTextString richTextString = creationHelper.createRichTextString(commentText);

		comment.setString(richTextString);
		comment.setAuthor("DataValidator");

		cell.setCellComment(comment);
    }

    public static void saveWorkBookChanges(Workbook workbook,String filePath)
    {
        try {
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
            fileOut.close();
            //workbook.close();
            logger.info("File written at :{}",filePath);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
    
    public static void writeExcelData(String sheetName,String filePath,List<List<String>> allDataList)
    {
 
        // Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
 
        // Creating a blank Excel sheet
        XSSFSheet sheet = workbook.createSheet(sheetName);
        int rownum = 0;
        for (List<String> rowList :allDataList) {
        	// Creating a new row in the sheet
            Row row = sheet.createRow(rownum++);
            int cellnum = 0;
            for(String data:rowList) 
            {
            	 // This line creates a cell in the next
                //  column of that row
                Cell cell = row.createCell(cellnum++);
 
                
                    cell.setCellValue(data);
            }
			
		}
 
       
 
        // Try block to check for exceptions
        try {
 
            // Writing the workbook
            FileOutputStream out = new FileOutputStream(
                new File(filePath));
            workbook.write(out);
 
            // Closing file output connections
            out.close();
 
            // Console message for successful execution of
            // program
            System.out.println(filePath+
                "  written successfully on disk.");
        }
 
        // Catch block to handle exceptions
        catch (Exception e) {
 
            // Display exceptions along with line number
            // using printStackTrace() method
            e.printStackTrace();
        }
    }
    
    public static void writeExcelData(String sheetName,String filePath)
    {
 
        // Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
 
        // Creating a blank Excel sheet
        XSSFSheet sheet
            = workbook.createSheet(sheetName);
 
        // Creating an empty TreeMap of string and Object][]
        // type
        Map<String, Object[]> data
            = new TreeMap<String, Object[]>();
 
        // Writing data to Object[]
        // using put() method
        data.put("1",
                 new Object[] { "ID", "NAME", "LASTNAME" });
        data.put("2",
                 new Object[] { 1, "Pankaj", "Kumar" });
        data.put("3",
                 new Object[] { 2, "Prakashni", "Yadav" });
        data.put("4", new Object[] { 3, "Ayan", "Mondal" });
        data.put("5", new Object[] { 4, "Virat", "kohli" });
 
        // Iterating over data and writing it to sheet
        Set<String> keyset = data.keySet();
 
        int rownum = 0;
 
        for (String key : keyset) {
 
            // Creating a new row in the sheet
            Row row = sheet.createRow(rownum++);
 
            Object[] objArr = data.get(key);
 
            int cellnum = 0;
 
            for (Object obj : objArr) {
 
                // This line creates a cell in the next
                //  column of that row
                Cell cell = row.createCell(cellnum++);
 
                if (obj instanceof String)
                    cell.setCellValue((String)obj);
 
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
 
        // Try block to check for exceptions
        try {
 
            // Writing the workbook
            FileOutputStream out = new FileOutputStream(
                new File(filePath));
            workbook.write(out);
 
            // Closing file output connections
            out.close();
 
            // Console message for successful execution of
            // program
            System.out.println(filePath+
                "  written successfully on disk.");
        }
 
        // Catch block to handle exceptions
        catch (Exception e) {
 
            // Display exceptions along with line number
            // using printStackTrace() method
            e.printStackTrace();
        }
    }
    public static void main(String[] args)
    {
        
        String filePath1 = "C://data-validator-test/mapping.xlsx";
        Workbook workbook = getWorkbookFromExcel(filePath1);
        Sheet sheet = getSheetFromWorkbook(workbook,0);
        List<String> columnList  = getColumnListFromSheet(sheet);
        List<String> columnLengthList  = getRowDataListFromSheet(sheet, 1);
        List<MappingCellItem> validationMappingDataList  = getValidationMappingDataListFromSheet(sheet);
        logger.info(""+validationMappingDataList);
        //Map<Integer, Map> dataset2 = readExcelSheetByName(filePath1,0,startRowNumber);
        Map<Integer, Map> dataMap = getDataFromSheetCustomized(sheet, columnList);
        logger.info("dataMap"+dataMap);
        
        List<LineSplitInfo> splitInfoList =StringUtil.getValidationMappingData(columnLengthList);
        logger.info("splitInfoList:"+splitInfoList);
        String textFileLocation = "C://data-validator-test//test.txt";
        
        List<List<String>> rawDataList =TextReader.getTextDataAsList(textFileLocation, splitInfoList);
        List<List<String>> totalDataList = new ArrayList<>();
        totalDataList.add(columnList);
        totalDataList.addAll(rawDataList);
        System.out.println(totalDataList);
        
        String outputExcelFileLocation = "C://data-validator-test//output.xlsx";
        String sheetName ="test";
        //writeExcelData(sheetName, outputExcelFileLocation);
        writeExcelData(sheetName, outputExcelFileLocation,totalDataList);



    }
}
