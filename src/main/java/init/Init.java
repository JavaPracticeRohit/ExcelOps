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

import commonUtils.ApplicationConstants;
import commonUtils.CommonUtilities;
import data.RowData;
import service.XlsService;

public class Init {
	private static final Logger LOGGER = Logger.getLogger(Init.class.getName());
	private static List<String> alErrors = new ArrayList<>();

	public static void main(String args[]) {

		Init objInitializer = new Init();
		objInitializer.initiate();
	}

	private void initiate() {

		XlsService objXlsService = new XlsService();

		String txFilePath = null;
		String aggrUserFilePath = null;
		String agr1251FilePath = null;
		String destinationPath = null;
		boolean fileCheck = false;
		boolean exit = false;
		// Scanner object`
		Scanner in = new Scanner(System.in);
		List<String> allowedTypes = CommonUtilities.getAllowedFileTypes();
		Map<Integer, String> txCodeSheetHeaderMap = CommonUtilities.prepareTxCodeHeaderMap();
		try {
			do {
				sendToConsole("-------------------------------------------------");
				sendToConsole(CommonUtilities.getDescription("welcometext1"));
				sendToConsole(CommonUtilities.getDescription("wtEntryId"));
				sendToConsole(CommonUtilities.getDescription("wtCount"));
				sendToConsole(CommonUtilities.getDescription("wtLuwCount"));
				sendToConsole(CommonUtilities.getDescription("wtGuiTime"));
				txFilePath = in.nextLine();
				sendToConsole(CommonUtilities.getDescription("wtAgrFilePath"));
				aggrUserFilePath = in.nextLine();
				sendToConsole(CommonUtilities.getDescription("wtAgr1251FilePath"));
				agr1251FilePath = in.nextLine();
				sendToConsole(CommonUtilities.getDescription("wtDestination"));
				destinationPath = in.nextLine();
				if (txFilePath == null || "".equals(txFilePath.trim())) {
					sendToConsole(CommonUtilities.getDescription("errDataFilePath"));
				}
				if (destinationPath == null || "".equals(destinationPath.trim())) {
					sendToConsole(CommonUtilities.getDescription("errDestination"));
				} else {
					isValidPath(txFilePath);
					isValidDir(destinationPath);
					if (alErrors.isEmpty()) {
						try {
							String fileType = getExtension(txFilePath);
							if (allowedTypes.contains(fileType)) {
								fileCheck = true;
								switch (fileType) {
								case ApplicationConstants.XLS:
									Map<String, RowData> txCodeDataMap = objXlsService.readAndConsolidateTxCodeData(txFilePath);
									Map<String, List<String>> aggrUserDataMap = objXlsService.readAndConsolidateAggrUserData(aggrUserFilePath);
									Map<String, RowData> agr1251DataMap = objXlsService.readAndConsolidateAggr1251UserData(agr1251FilePath, aggrUserDataMap,txCodeDataMap);
									HSSFWorkbook workbook;
									// name of excel file
									String excelFileName = System.currentTimeMillis() + "";
									File file = new File(destinationPath + File.separator + excelFileName + "."+ApplicationConstants.XLS);
									if (!file.exists()) {
										// Create new file if it does not exist
										workbook = new HSSFWorkbook();
									} else {
										try (
												// Make current input to exist file
												InputStream is = new FileInputStream(file)) {
											workbook = new HSSFWorkbook(is);
										}
									}
									objXlsService.prepareExcel(txCodeSheetHeaderMap, agr1251DataMap, workbook);
									FileOutputStream fileOut = new FileOutputStream(file);
									// write this workbook to an Outputstream.
									workbook.write(fileOut);
									fileOut.flush();
									fileOut.close();
									Desktop.getDesktop().open(file);
									break;
								default:
									sendToConsole(CommonUtilities.getDescription("errInvalidFileType",fileType));
								}
							} else {
								sendToConsole(CommonUtilities.getDescription("errInvalidFileType"));
							}
						} catch (Exception e) {
							alErrors.add(e.getMessage());
						}
					} else {
						printErr();
					}
				}
			} while (!fileCheck);
			do {
				printErr();
				sendToConsole(CommonUtilities.getDescription("exitText"));
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
		alErrors = new ArrayList<>();
	}

	private void sendToConsole(String message) {
		System.out.println(message);
	}

	private boolean isValidPath(String path) {
		File file = new File(path);
		if (file.exists()) {
			if (file.isFile()) {
				return true;
			} else if (file.isDirectory()) {
				alErrors.add(CommonUtilities.getDescription("errCompFilePath") + path);
				return false;
			}
		}
		alErrors.add(CommonUtilities.getDescription("errInvalidPath") + path);
		return false;
	}

	private boolean isValidDir(String path) {
		File file = new File(path);
		if (file.isDirectory()) {
			return true;
		}
		alErrors.add(CommonUtilities.getDescription("errInvalidDir") + path);
		return false;
	}

	private String getExtension(String fileName) {
		char ch;
		int len;
		if (fileName == null || (len = fileName.length()) == 0 || (ch = fileName.charAt(len - 1)) == '/' || ch == '\\'
				|| // in the case of a directory
				ch == '.') // in the case of . or ..
			return "";
		int dotInd = fileName.lastIndexOf('.');
		int sepInd = Math.max(fileName.lastIndexOf('/'), fileName.lastIndexOf('\\'));
		if (dotInd <= sepInd)
			return "";
		else
			return fileName.substring(dotInd + 1).toLowerCase();
	}
}
