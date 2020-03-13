/**
 * Code created by Rohit Bhatia for self use or Demo purpose only.
 */
package service;

import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import commonUtils.ApplicationConstants;
import data.RowData;
import errorHandling.ColumnNameNotFound;
import xlsUtils.XlsReaderUtils;
import xlsUtils.XlsWriterUtils;

public class XlsService {

	private static XlsReaderUtils objXlsReader = new XlsReaderUtils(); // XLS Reader object
	private static XlsWriterUtils objXlsWriter = new XlsWriterUtils(); // XLS Writer object

	public Map<String, RowData> readAndConsolidateTxCodeData(String inputXlsPath) throws IOException, ColumnNameNotFound {
		String groupByColumn = ApplicationConstants.ENTRY_ID;
		Map<String, Integer> headerColIdxMap = new HashMap<>();
		Map<String, RowData> dataMap;
		HSSFWorkbook wb = objXlsReader.getXlsWorkbook(inputXlsPath);
		int sheetNo = 0; // Sheet To read
		int headerRowNum = 0;
		// Set the indexes of required columns
		mapColIndex(headerColIdxMap, wb, sheetNo, headerRowNum, ApplicationConstants.ENTRY_ID);
		mapColIndex(headerColIdxMap, wb, sheetNo, headerRowNum, ApplicationConstants.COUNT);
		mapColIndex(headerColIdxMap, wb, sheetNo, headerRowNum, ApplicationConstants.LUW_COUNT);
		mapColIndex(headerColIdxMap, wb, sheetNo, headerRowNum, ApplicationConstants.GUITIME);

		Map<Integer, String> colNameIdxMap = createIdxColMap(headerColIdxMap);

		dataMap = objXlsReader.getColVals(wb, sheetNo, headerColIdxMap, colNameIdxMap, groupByColumn);
		return dataMap;

	}
	
	public Map<String, RowData> readAndConsolidateAggr1251UserData(String inputFilePath, Map<String, List<String>> aggrUserDataMap, Map<String, RowData> txCodeDataMap) throws IOException, ColumnNameNotFound  {
		HSSFWorkbook wb = objXlsReader.getXlsWorkbook(inputFilePath);
		int sheetNo = 0; // Sheet To read
		return objXlsReader.readAgr1251Sheet(wb, sheetNo,aggrUserDataMap,txCodeDataMap);
	}

	public Map<String, List<String>> readAndConsolidateAggrUserData(String aggrUserFilePath) throws Exception {
		HSSFWorkbook wb = objXlsReader.getXlsWorkbook(aggrUserFilePath);
		int sheetNo = 0; // Sheet To read
		return objXlsReader.readAgrSheet(wb, sheetNo);
	}

	private Map<Integer, String> createIdxColMap(Map<String, Integer> headerColIdxMap) {
		Map<Integer, String> colNameIdxMap = new HashMap<>();
		Iterator<String> itr = headerColIdxMap.keySet().iterator();
		while (itr.hasNext()) {
			String colName = itr.next();
			int colIdx = headerColIdxMap.get(colName);
			colNameIdxMap.put(colIdx, colName);
		}
		return colNameIdxMap;
	}

	private List<String> getDistinctValsForCol(HSSFWorkbook wb, int sheetNo, Integer colIndex) {
		List<String> vals = objXlsReader.getAllValsForCol(wb, sheetNo, colIndex);
		return vals.stream().distinct().collect(Collectors.toList());
	}

	private void mapColIndex(Map<String, Integer> headerColIdxMap, HSSFWorkbook wb, int sheetNo, int headerRowNum,
			String matchColName) throws ColumnNameNotFound {
		int matchedColNum = objXlsReader.getColNoForColName(wb, sheetNo, headerRowNum, matchColName);
		if (matchedColNum < 0)
			throw new ColumnNameNotFound("Column Name [" + matchColName + "] not found in sheet");
		headerColIdxMap.put(matchColName, matchedColNum);
	}

	public HSSFWorkbook prepareExcel(Map<Integer, String> headerMap, Map<String, RowData> dataMap, HSSFWorkbook workbook) throws IOException {
		String sheetName = "Sheet1";// name of sheet
		objXlsWriter.writeHeadersToXls(workbook, sheetName, headerMap, 0);
		objXlsWriter.writeDataToXls(workbook, sheetName,dataMap);
		return workbook;
	}

	

}
