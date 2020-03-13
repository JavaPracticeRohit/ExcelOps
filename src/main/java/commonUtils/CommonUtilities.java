/**
 * 
 */
package commonUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;

/**
 * This class contains Common Methods
 *
 */
public class CommonUtilities {

	/**
	 * Method returns the map representation of the columns expected in Tx Code
	 * sheet.
	 * 
	 * @return Map : Key - Column Index : Value : Column Name
	 */
	public static Map<Integer, String> prepareTxCodeHeaderMap() {
		Map<Integer, String> txCodeSheetHeaderMap = new HashMap<>();
		txCodeSheetHeaderMap.put(0, ApplicationConstants.TRANSACTION_CODE);
		txCodeSheetHeaderMap.put(1, ApplicationConstants.SUM_TX_COUNT);
		txCodeSheetHeaderMap.put(2, ApplicationConstants.SUM_STEP_COUNT);
		txCodeSheetHeaderMap.put(3, ApplicationConstants.SUM_GUI_TIME);
		txCodeSheetHeaderMap.put(4, ApplicationConstants.ROLES);
		txCodeSheetHeaderMap.put(5, ApplicationConstants.USERS);
		return txCodeSheetHeaderMap;
	}

	/**
	 * Method returns the list of file types accepted by the application as
	 * input.
	 * 
	 * @return List containing the accepted file extensions
	 */
	public static List<String> getAllowedFileTypes() {
		List<String> allowedTypes = new ArrayList<>();
		allowedTypes.add(ApplicationConstants.XLS);
		allowedTypes.add(ApplicationConstants.XLSX);
		return allowedTypes;
	}

	/**
	 * Constructor
	 */
	private CommonUtilities() {

	}
	
	  /**
     * Method returns the complete description from the property file (based on
     * locale) for the specific abbreviation key.
     *
     * @param abbr
     *            : abbreviation key
     * @return description based on locale
     */
    public static String getDescription(String abbr) {
        if (abbr == null || abbr.trim().equals(ApplicationConstants.BLANK))
            return ApplicationConstants.BLANK;
        ResourceBundle bundle = ResourceBundle.getBundle("resources.appl");
        return bundle.getString(abbr);
    }

 

    /**
     * Method returns the complete description from the property file (based on
     * locale) for the specific abbreviation key.
     *
     * @param abbr
     *            : abbreviation key
     * @param val1
     *            : this value will replace the %1 in the description
     * @return description based on locale
     */
    public static String getDescription(String abbr, String val1) {
        String description = getDescription(abbr);
        if (description == null || description.trim().equals(ApplicationConstants.BLANK))
            return ApplicationConstants.BLANK;
        description = description.replace("%1", val1);
        return description;
    }

 

    /**
     * Method returns the complete description from the property file (based on
     * locale) for the specific abbreviation key.
     *
     * @param abbr
     *            : abbreviation key
     * @param val1
     *            : this value will replace the %1 in the description
     * @param val2
     *            : this value will replace the %2 in the description
     * @return description based on locale
     */
    public static String getDescription(String abbr, String val1, String val2) {
        String description = getDescription(abbr);
        if (description == null || description.trim().equals(ApplicationConstants.BLANK))
            return ApplicationConstants.BLANK;
        description = description.replace("%1", val1).replace("%2", val2);
        return description;
    }

 

    /**
     * Method returns the complete description from the property file (based on
     * locale) for the specific abbreviation key.
     *
     * @param abbr
     *            : abbreviation key
     * @param val1
     *            : this value will replace the %1 in the description
     * @param val2
     *            : this value will replace the %2 in the description
     * @param val3
     *            : this value will replace the %3 in the description
     * @return description based on locale
     */
    public static String getDescription(String abbr, String val1, String val2, String val3) {
        String description = getDescription(abbr);
        if (description == null || description.trim().equals(ApplicationConstants.BLANK))
            return ApplicationConstants.BLANK;
        description = description.replace("%1", val1).replace("%2", val2).replace("%3", val3);
        return description;
    }
}
