package texttoexcelconverter.application;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import texttoexcelconverter.model.LineSplitInfo;
import texttoexcelconverter.model.MappingCellItem;
import texttoexcelconverter.util.ExcelUtil;
import texttoexcelconverter.util.StringUtil;
import texttoexcelconverter.util.TextReader;

import java.lang.invoke.MethodHandles;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ApplicationWithParamaters
{
    public static final Logger logger = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass());
    public static void main(String[] args) {
            convertTextToExcel(args[0],args[1],args[2]);

    }

    public static void convertTextToExcel(String mappingFilePath,String textFileLocation,String outputFilePath) {
        Workbook workbook = ExcelUtil.getWorkbookFromExcel(mappingFilePath);
        Sheet sheet = ExcelUtil.getSheetFromWorkbook(workbook,0);
        List<String> columnList  = ExcelUtil.getColumnListFromSheet(sheet);
        List<String> columnLengthList  = ExcelUtil.getRowDataListFromSheet(sheet, 1);
        List<MappingCellItem> validationMappingDataList  = ExcelUtil.getValidationMappingDataListFromSheet(sheet);
        logger.info(""+validationMappingDataList);
        //Map<Integer, Map> dataset2 = readExcelSheetByName(filePath1,0,startRowNumber);
        Map<Integer, Map> dataMap = ExcelUtil.getDataFromSheetCustomized(sheet, columnList);
        logger.info("dataMap"+dataMap);

        List<LineSplitInfo> splitInfoList = StringUtil.getValidationMappingData(columnLengthList);
        logger.info("splitInfoList:"+splitInfoList);

        List<List<String>> rawDataList = TextReader.getTextDataAsList(textFileLocation, splitInfoList);
        List<List<String>> totalDataList = new ArrayList<>();
        totalDataList.add(columnList);
        totalDataList.addAll(rawDataList);
        System.out.println(totalDataList);

        String sheetName ="test";
        //writeExcelData(sheetName, outputExcelFileLocation);
        ExcelUtil.writeExcelData(sheetName, outputFilePath,totalDataList);



    }
}
