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

import data.RowData;

/**
 * @author blchi
 *
 */
public class XlsReaderUtils {

	/**
	 * Method to get the workbook object from xls file
	 * @param inputFile
	 * @return HSSFWorkbook Object
	 * @throws IOException 
	 */
	public HSSFWorkbook getXlsWorkbook(String inputFile) throws IOException {
		InputStream ExcelFileToRead = new FileInputStream(inputFile);
		HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
		return wb;
	}

	/**
	 * Method to get all the entries in a column
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

	public List<String> getAllValsForCol(HSSFWorkbook wb, int sheetNo, int colIndex) {
		List<String> alVals = new ArrayList<String>();
		HSSFSheet sheet = wb.getSheetAt(sheetNo);
		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			HSSFRow row = sheet.getRow(rowIdx);
			HSSFCell cell = row.getCell(colIndex);
			alVals.add(cell.getStringCellValue());
		}
		return alVals;
	}

	public Map<String, List<String>> readAgrSheet(HSSFWorkbook wb, int sheetNo) throws Exception {
		Map<String, List<String>> hmRowData = new HashMap<String, List<String>>();
		Map<String, Integer> hmColIdxMap = new LinkedHashMap<String, Integer>();
		List<String> users = new ArrayList<>();
		HSSFSheet sheet = wb.getSheetAt(sheetNo);
		HSSFRow headerRow = sheet.getRow(0);
		Iterator cells = headerRow.cellIterator();
		while (cells.hasNext()) {
			HSSFCell headerCell = (HSSFCell) cells.next();
			hmColIdxMap.put(headerCell.getStringCellValue(), headerCell.getColumnIndex());
		}

		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			HSSFRow row = sheet.getRow(rowIdx);
			String role = getStringVal(hmColIdxMap, row, "AGR_NAME");
			String user = getStringVal(hmColIdxMap, row, "UNAME");
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

	public Map<String, RowData> readAgr1251Sheet(HSSFWorkbook wb, int sheetNo,
			Map<String, List<String>> aggrUserDataMap, Map<String, RowData> txCodeDataMap) throws Exception {
		Map<String, RowData> hmRowData = new HashMap<String, RowData>();
		Map<String, Integer> hmColIdxMap = new LinkedHashMap<String, Integer>();
		RowData objRowData = new RowData();
		List<String> roles = new ArrayList<>();
		HSSFSheet sheet = wb.getSheetAt(sheetNo);
		HSSFRow headerRow = sheet.getRow(0);
		Iterator cells = headerRow.cellIterator();
		while (cells.hasNext()) {
			HSSFCell headerCell = (HSSFCell) cells.next();
			hmColIdxMap.put(headerCell.getStringCellValue(), headerCell.getColumnIndex());
		}

		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			HSSFRow row = sheet.getRow(rowIdx);
			String role = getStringVal(hmColIdxMap, row, "AGR_NAME");
			String txId = getStringVal(hmColIdxMap, row, "LOW");
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
				roles = new ArrayList<>();
			}
		}

		return hmRowData;
	}

	private String getStringVal(Map<String, Integer> hmColIdxMap, HSSFRow row, String key) throws Exception {
		HSSFCell cell;
		String val = null;
		if (!hmColIdxMap.containsKey(key))
			throw new Exception("Column" + key + " not found in sheet");
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

	public Map<String, RowData> getColVals(HSSFWorkbook wb, int sheetNo, Map<String, Integer> headerColIdxMap,
			Map<Integer, String> colNameIdxMap, String groupByColumn) {
		HSSFCell cell;
		Map<String, RowData> dataMap = new HashMap<String, RowData>();
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
							bgCellVal = new BigDecimal(cell.getNumericCellValue());
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
