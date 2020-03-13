/**
 * Code created by Rohit Bhatia for self use or Demo purpose only.
 */
package xlsUtils;

import java.io.IOException;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import data.RowData;

public class XlsWriterUtils {

	public HSSFWorkbook writeHeadersToXls(HSSFWorkbook wb, String sheetName, Map<Integer, String> dataMap, int rowNum)
			throws IOException {
		HSSFSheet sheet = null;
		// Check if the workbook is empty or not
		if (wb.getNumberOfSheets() != 0) {
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				if (wb.getSheetName(i).equals(sheetName)) {
					sheet = wb.getSheet(sheetName);
				} else
					sheet = wb.createSheet(sheetName);
			}
		} else {
			// Create new sheet to the workbook if empty
			sheet = wb.createSheet(sheetName);
		}

		HSSFRow row = sheet.createRow(rowNum);
		Iterator<Integer> itrRowData = dataMap.keySet().iterator();
		while (itrRowData.hasNext()) {
			int key = itrRowData.next();
			String val = dataMap.get(key);
			HSSFCell cell = row.createCell(key);
			cell.setCellValue(val);
		}
		return wb;
	}
	
	public HSSFWorkbook writeDataToXls(HSSFWorkbook wb, String sheetName, Map<String, RowData> dataMap) throws IOException {
		HSSFSheet sheet = wb.getSheet(sheetName);
		int rowNum = 1;
		Iterator<String> itrDataMap = dataMap.keySet().iterator();
		while (itrDataMap.hasNext()) {
			String key = itrDataMap.next();
			RowData objRow = dataMap.get(key);
			HSSFRow row = sheet.createRow(rowNum);

			HSSFCell txCodeCell = row.createCell(0);
			txCodeCell.setCellType(CellType.STRING);
			txCodeCell.setCellValue(objRow.getTxCode());

			HSSFCell txCountCell = row.createCell(1);
			txCountCell.setCellType(CellType.NUMERIC);
			txCountCell.setCellValue(objRow.getBgSumTxCount().toString());

			HSSFCell txStepCountCell = row.createCell(2);
			txStepCountCell.setCellType(CellType.NUMERIC);
			txStepCountCell.setCellValue(objRow.getBgSumStepCount().toString());

			HSSFCell txGuiTimeCell = row.createCell(3);
			txGuiTimeCell.setCellType(CellType.NUMERIC);
			txGuiTimeCell.setCellValue(objRow.getBgSumGuiTime().toString());

			HSSFCell txRoles = row.createCell(4);
			txRoles.setCellType(CellType.NUMERIC);
			txRoles.setCellValue(objRow.getRoles().size());

			HSSFCell txUsers = row.createCell(5);
			txUsers.setCellType(CellType.NUMERIC);
			txUsers.setCellValue(objRow.getUsers().size());
			rowNum++;
		}

		return wb;
	}
	
}
