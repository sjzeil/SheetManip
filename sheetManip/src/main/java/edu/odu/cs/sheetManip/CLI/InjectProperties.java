package edu.odu.cs.sheetManip.CLI;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import edu.odu.cs.sheetManip.Spreadsheet;


/**
 * Utility for populating a sheet from a CSV file.
 * 
 * Usage: java edu.odu.cs.gradeSync.CLI.InjectProperties spreadsheetFile propertiesFile leftDelim rightDelim
 * 
 *     Delimiters default to {{ and }}
 * 
 * @author zeil
 *
 */
public class InjectProperties {


	private static Log log = LogFactory.getLog(InjectProperties.class);

	private String spreadsheetFileName;
	private  String propertiesFileName;
	private String leftDelimeter;
    private String rightDelimeter;


	/**
	 * Launch the application.
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		if (args.length < 2 || args.length > 4) {
			log.error("Usage: java edu.odu.cs.gradeSync.CLI.InjectProperties spreadsheetFile propertiesFile leftDelim rightDelim");
			System.exit(1);
		}
		String arg2 = (args.length >= 3) ? args[2] : "{{";
        String arg3 = (args.length >= 4) ? args[3] : "}}";
		new InjectProperties(args[0], args[1], arg2, arg3).run();
	}

	/**
	 * Create the application.
	 * 
	 * @param spreadsheetFileName   path to an Excel spreadsheet (.xls or .xlsx)
	 * @param propertiesFileName    path to a java.util.Properties file
	 * @param leftDelim             string appearing to the left of each
	 *                                 embedded property name
     * @param rightDelim            string appearing to the right of each
     *                                 embedded property name
	 */
	public InjectProperties(String spreadsheetFileName, String propertiesFileName,
	        String leftDelim, String rightDelim) {
		
		this.spreadsheetFileName = spreadsheetFileName;
		this.propertiesFileName = propertiesFileName;
		this.leftDelimeter = leftDelim;
		this.rightDelimeter = rightDelim;
	}

	public void run() throws IOException {
		File ssFile = new File(spreadsheetFileName);
		File propFile = new File(propertiesFileName);
		BufferedReader reader = null;
		Spreadsheet ss = null;
		try {
		    ss = new Spreadsheet(ssFile);
		    Properties properties = new Properties();
		    reader = new BufferedReader(new FileReader(propFile));
		    properties.load(reader);
		    ss.injectProperties(properties, leftDelimeter, rightDelimeter);
		} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
			log.error ("Unable to load sheet from temporary file", e);
			return;
		} finally {
		    if (reader != null)
		        reader.close();
		    if (ss != null)
		        ss.close();
		}
	}


}
