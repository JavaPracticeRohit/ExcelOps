package init;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import data.RowData;
import service.XlsService;

public class Init {
	private static final Logger LOGGER = Logger.getLogger(Init.class.getName());
	private static List<String> alErrors = new ArrayList<String>();

	public static void main(String args[]) throws IOException {

		Init objInitializer = new Init();
		objInitializer.initiate();
	}

	private void initiate() {
		
		XlsService objXlsService = new XlsService();
		
		String txFilePath = null;
		String aggrUserFilePath = null;
		String agr_1251FilePath = null;
		String destinationPath = null;
		Scanner in = new Scanner(System.in);
		boolean fileCheck = false;
		boolean exit = false;
		List<String> allowedTypes = new ArrayList<String>();
		allowedTypes.add("xls");
		allowedTypes.add("xlsx");
		Map<Integer,String> txCodeSheetHeaderMap = new HashMap<Integer, String>();
		txCodeSheetHeaderMap.put(0, "Transaction code");
		txCodeSheetHeaderMap.put(1, "Sum of Tx Count");
		txCodeSheetHeaderMap.put(2, "Sum of Steps Count");
		txCodeSheetHeaderMap.put(3, "Sum of GUITIME");
		txCodeSheetHeaderMap.put(4, "Roles");
		txCodeSheetHeaderMap.put(5, "Users");
		try {
			do {
				sendToConsole("-------------------------------------------------");
				sendToConsole(
						"Full path(including file name eg:[c:\\<folderName>\\<folderName>\\test.xls]) for Xls Data file containing the following column as follows : ");
				sendToConsole("1. ENTRY_ID -> Will be termed as Tx Code");
				sendToConsole("2. Count");
				sendToConsole("3. LUW_Count");
				sendToConsole("4. GUI Time");
				txFilePath = in.nextLine();
				sendToConsole(
						"Full path(including file name eg:[c:\\<folderName>\\<folderName>\\test.xls]) for Xls Data file containing the AGR_USERS Data : ");
				aggrUserFilePath = in.nextLine();
				sendToConsole(
						"Full path(including file name eg:[c:\\<folderName>\\<folderName>\\test.xls]) for Xls Data file containing the AGR_1251 Data : ");
				agr_1251FilePath = in.nextLine();
				sendToConsole("Destination Folder/Directory Path : ");
				destinationPath = in.nextLine();
				if (txFilePath == null || "".equals(txFilePath.trim())) {
					sendToConsole("Data File Path : Not provided ");
				}
				if (destinationPath == null || "".equals(destinationPath.trim())) {
					sendToConsole("Destination File Path : Not provided ");
				} else {
					isValidPath(txFilePath);
					isValidDir(destinationPath);
					if (alErrors.isEmpty()) {
						try{
						String fileType = getExtension(txFilePath);
						if (allowedTypes.contains(fileType)) {
							fileCheck = true;
							switch (fileType) {
							case "xls":	
								Map<String, RowData> txCodeDataMap = objXlsService.readAndConsolidateTxCodeData(txFilePath);
								Map<String, List<String>> aggrUserDataMap = objXlsService.readAndConsolidateAggrUserData(aggrUserFilePath);
								Map<String, RowData> agr_1251DataMap = objXlsService.readAndConsolidateAggr1251UserData(agr_1251FilePath,aggrUserDataMap,txCodeDataMap);
								
								
								HSSFWorkbook workbook;
								String excelFileName = System.currentTimeMillis() + "";// name of excel file
								File file = new File(destinationPath + "/" + excelFileName + ".xls");
								if (file.exists() == false) {
									// Create new file if it does not exist
									workbook = new HSSFWorkbook();
								} else {
									try (
											// Make current input to exist file
											InputStream is = new FileInputStream(file)) {
										workbook = new HSSFWorkbook(is);
									}
								}
								
								objXlsService.prepareExcel(destinationPath,txCodeSheetHeaderMap,agr_1251DataMap,workbook);
								FileOutputStream fileOut = new FileOutputStream(file);
								// write this workbook to an Outputstream.
								workbook.write(fileOut);
								fileOut.flush();
								fileOut.close();
								
								
								Desktop.getDesktop().open(file);
								break;
							default:
								sendToConsole("Implementation missing for " + fileType
										+ ". Kindly connect with technical Team.");
							}
						} else {
							sendToConsole("Invalid file type.");
						}
						}catch (Exception e) {
							e.printStackTrace();
							alErrors.add(e.getMessage());
						}
					} else {
						printErr();
					}
				}
			} while (!fileCheck);
			do {
				printErr();
				sendToConsole("Press x to exit...");
				String eCheck = in.nextLine();
				if ("x".equalsIgnoreCase(eCheck))
					exit = true;
			} while (!exit);
		} finally {
			in.close();
		}
	}

	private void printErr() {
		for (String err : alErrors) {
			sendToConsole(err);
		}
		alErrors = new ArrayList<String>();
	}

	private void sendToConsole(String message) {
		System.out.println(message);
		// LOGGER.info(message);
	}

	private boolean isValidPath(String path) {
		File file = new File(path);
		if (file.exists()) {
			if (file.isFile()) {
				return true;
			} else if (file.isDirectory()) {
				alErrors.add("Complete Path with File Name is needed  : " + path);
				return false;
			}
		}
		alErrors.add("Invalid path : " + path);
		return false;
	}

	private boolean isValidDir(String path) {
		File file = new File(path);
		if (file != null && file.isDirectory()) {
			return true;
		}
		alErrors.add("Invalid directory path  : " + path);
		return false;
	}

	private String getExtension(String fileName) {
		char ch;
		int len;
		if (fileName == null || (len = fileName.length()) == 0 || (ch = fileName.charAt(len - 1)) == '/' || ch == '\\'
				|| //in the case of a directory
				ch == '.') //in the case of . or ..
			return "";
		int dotInd = fileName.lastIndexOf('.'),
				sepInd = Math.max(fileName.lastIndexOf('/'), fileName.lastIndexOf('\\'));
		if (dotInd <= sepInd)
			return "";
		else
			return fileName.substring(dotInd + 1).toLowerCase();
	}
}
