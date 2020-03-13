/**
 * Code created by Rohit Bhatia for self use or Demo purpose only.
 */
package xlsUtils;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import commonUtils.ApplicationConstants;
import commonUtils.CommonUtilities;
import data.RowData;
import errorHandling.ColumnNameNotFound;

public class XlsReaderUtils {

	/**
	 * Method to get the workbook object from xls file
	 * 
	 * @param inputFile
	 * @return HSSFWorkbook Object
	 * @throws IOException
	 */
	public HSSFWorkbook getXlsWorkbook(String inputFile) throws IOException {
		InputStream ExcelFileToRead = new FileInputStream(inputFile);
		return new HSSFWorkbook(ExcelFileToRead);
	}

	/**
	 * Method to get all the entries in a column
	 * 
	 * @param wb
	 * @param sheetNo
	 * @param colName
	 * @return
	 */
	public int getColNoForColName(HSSFWorkbook wb, int sheetNo, int headerRowNum, String colName) {
		int matchedColNum = -1;
		HSSFSheet sheet = wb.getSheetAt(sheetNo);
		HSSFRow headerRow = sheet.getRow(headerRowNum);
		Iterator cells = headerRow.cellIterator();
		while (cells.hasNext()) {
			HSSFCell cell = (HSSFCell) cells.next();
			if (colName.equalsIgnoreCase(cell.getStringCellValue())) {
				matchedColNum = cell.getColumnIndex();
				break;
			}
		}
		return matchedColNum;
	}

	/**
	 * Method to return all the values of specified column index in specified
	 * sheet no
	 * 
	 * @param wb
	 *            Workbook object
	 * @param sheetNo
	 *            Sheet No
	 * @param colIndex
	 *            Column Index
	 * @return Column values in a list
	 */
	public List<String> getAllValsForCol(HSSFWorkbook wb, int sheetNo, int colIndex) {
		List<String> alVals = new ArrayList<>();
		HSSFSheet sheet = wb.getSheetAt(sheetNo);
		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			HSSFRow row = sheet.getRow(rowIdx);
			HSSFCell cell = row.getCell(colIndex);
			alVals.add(cell.getStringCellValue());
		}
		return alVals;
	}

	/**
	 * Method to read all rows in specified sheet
	 * 
	 * @param wb
	 *            workbook object
	 * @param sheetNo
	 * @return rows in a map
	 * @throws ColumnNameNotFound
	 *             is thrown if the expected column is not found in the sheet
	 */
	public Map<String, List<String>> readAgrSheet(HSSFWorkbook wb, int sheetNo) throws ColumnNameNotFound {
		Map<String, List<String>> hmRowData = new HashMap<>();
		Map<String, Integer> hmColIdxMap = new LinkedHashMap<>();
		List<String> users;
		HSSFSheet sheet = wb.getSheetAt(sheetNo);
		HSSFRow headerRow = sheet.getRow(0);
		Iterator cells = headerRow.cellIterator();
		while (cells.hasNext()) {
			HSSFCell headerCell = (HSSFCell) cells.next();
			hmColIdxMap.put(headerCell.getStringCellValue(), headerCell.getColumnIndex());
		}

		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			HSSFRow row = sheet.getRow(rowIdx);
			String role = getStringVal(hmColIdxMap, row, ApplicationConstants.AGR_NAME);
			String user = getStringVal(hmColIdxMap, row, ApplicationConstants.UNAME);
			if (role != null && user != null) {
				if (hmRowData.containsKey(role))
					users = hmRowData.get(role);
				else
					users = new ArrayList<>();
				if (!users.contains(user))
					users.add(user);
				hmRowData.put(role, users);
			}
		}

		return hmRowData;
	}

	/**
	 * Method to read all rows in specified sheet and consolidate it on the
	 * basis of values already present in the maps
	 * 
	 * @param wb
	 *            workbook object
	 * @param sheetNo
	 *            sheet no
	 * @param aggrUserDataMap
	 *            Map(AgrUserData) against which the values are to be
	 *            consolidated
	 * @param txCodeDataMap
	 *            Map(Transaction Count) against which the values are to be
	 *            consolidated
	 * @return Map containing the consolidated values
	 * @throws ColumnNameNotFound
	 *             is thrown if the expected column is not found in the sheet
	 */
	public Map<String, RowData> readAgr1251Sheet(HSSFWorkbook wb, int sheetNo,
			Map<String, List<String>> aggrUserDataMap, Map<String, RowData> txCodeDataMap) throws ColumnNameNotFound {
		Map<String, RowData> hmRowData = new HashMap<>();
		Map<String, Integer> hmColIdxMap = new LinkedHashMap<>();
		RowData objRowData;
		List<String> roles;
		HSSFSheet sheet = wb.getSheetAt(sheetNo);
		HSSFRow headerRow = sheet.getRow(0);
		Iterator cells = headerRow.cellIterator();
		while (cells.hasNext()) {
			HSSFCell headerCell = (HSSFCell) cells.next();
			hmColIdxMap.put(headerCell.getStringCellValue(), headerCell.getColumnIndex());
		}

		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			HSSFRow row = sheet.getRow(rowIdx);
			String role = getStringVal(hmColIdxMap, row, ApplicationConstants.AGR_NAME);
			String txId = getStringVal(hmColIdxMap, row, ApplicationConstants.LOW);
			if (txId != null && role != null) {
				if (hmRowData.containsKey(txId))
					objRowData = hmRowData.get(txId);
				else
					objRowData = new RowData();
				objRowData.setTxCode(txId);
				roles = objRowData.getRoles();
				roles.add(role);
				objRowData.setRoles(roles);

				if (aggrUserDataMap.containsKey(role)) {
					objRowData.setUsers(aggrUserDataMap.get(role));
				} else {
					objRowData.setUsers(new ArrayList<>());
				}
				if (txCodeDataMap.containsKey(txId)) {
					RowData rowDataTxCodeSheet = txCodeDataMap.get(txId);
					objRowData.setBgSumGuiTime(rowDataTxCodeSheet.getBgSumGuiTime());
					objRowData.setBgSumStepCount(rowDataTxCodeSheet.getBgSumStepCount());
					objRowData.setBgSumTxCount(rowDataTxCodeSheet.getBgSumTxCount());
				} else {
					objRowData.setBgSumGuiTime(new BigDecimal(0));
					objRowData.setBgSumStepCount(new BigDecimal(0));
					objRowData.setBgSumTxCount(new BigDecimal(0));
				}
				hmRowData.put(txId, objRowData);
			}
		}

		return hmRowData;
	}

	/**
	 * Method to return a particular column value against specific column name
	 * for specific excel row
	 * 
	 * @param hmColIdxMap
	 *            Map containing the column indexes
	 * @param row
	 *            Excel row containing the data
	 * @param key
	 *            Column Name
	 * @return Column Value
	 * @throws ColumnNameNotFound
	 *             is thrown if the expected column is not found in the sheet
	 */
	private String getStringVal(Map<String, Integer> hmColIdxMap, HSSFRow row, String key) throws ColumnNameNotFound {
		HSSFCell cell;
		String val = null;
		if (!hmColIdxMap.containsKey(key))
			throw new ColumnNameNotFound(CommonUtilities.getDescription("errColKey", key));
		cell = row.getCell(hmColIdxMap.get(key));
		if (cell == null)
			return val;

		switch (cell.getCellType()) {
		case NUMERIC:
			val = cell.getNumericCellValue() + "";
			break;
		case STRING:
			val = cell.getStringCellValue();
			break;
		}
		return val;
	}

	/**
	 * Method to return the consolidated values of specific column
	 * 
	 * @param wb
	 *            workbook object
	 * @param sheetNo
	 *            Sheet No
	 * @param headerColIdxMap
	 *            Column index map
	 * @param colNameIdxMap
	 *            Column Name Map
	 * @param groupByColumn
	 *            Column against which the data has to be consolidated
	 * @return consolidated values in map form
	 */
	public Map<String, RowData> getColVals(HSSFWorkbook wb, int sheetNo, Map<String, Integer> headerColIdxMap,
			Map<Integer, String> colNameIdxMap, String groupByColumn) {
		HSSFCell cell;
		Map<String, RowData> dataMap = new HashMap<>();
		HSSFSheet sheet = wb.getSheetAt(sheetNo);
		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			HSSFRow row = sheet.getRow(rowIdx);
			Iterator cells = row.cellIterator();
			while (cells.hasNext()) {
				cell = (HSSFCell) cells.next();
				int colIdx = cell.getColumnIndex();
				if (colNameIdxMap.containsKey(colIdx)) {
					RowData rowData = null;
					if (colIdx != headerColIdxMap.get(groupByColumn)) {
						HSSFCell txCell = row.getCell(headerColIdxMap.get(groupByColumn));
						String txVal = txCell.getStringCellValue();
						if (dataMap.containsKey(txVal))
							rowData = dataMap.get(txVal);
						else {
							rowData = new RowData();
							rowData.setTxCode(txVal);
						}
						BigDecimal bgCellVal = null;
						switch (cell.getCellType()) {
						case NUMERIC:
							bgCellVal = BigDecimal.valueOf(cell.getNumericCellValue());
							break;
						case STRING:
							bgCellVal = new BigDecimal(cell.getStringCellValue());
							break;
						default:
							break;

						}

						switch (colIdx) {
						case 3:
							BigDecimal bgTxCount = rowData.getBgSumTxCount();
							bgTxCount = bgTxCount.add(bgCellVal != null ? bgCellVal : new BigDecimal(0));
							rowData.setBgSumTxCount(bgTxCount);
							break;
						case 4:
							BigDecimal bgSumStepCount = rowData.getBgSumStepCount();
							bgSumStepCount = bgSumStepCount.add(bgCellVal != null ? bgCellVal : new BigDecimal(0));
							rowData.setBgSumStepCount(bgSumStepCount);
							break;
						case 10:
							BigDecimal bgSumGuiTime = rowData.getBgSumGuiTime();
							bgSumGuiTime = bgSumGuiTime.add(bgCellVal != null ? bgCellVal : new BigDecimal(0));
							rowData.setBgSumGuiTime(bgSumGuiTime);
							break;
						default:
							break;
						}

						dataMap.put(rowData.getTxCode(), rowData);
					}

				}
			}
		}
		return dataMap;
	}

}
