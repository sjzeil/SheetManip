package edu.odu.cs.sheetManip.CLI;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.opencsv.CSVReader;

import edu.odu.cs.sheetManip.Spreadsheet;


/**
 * Utility for extracting a sheet as a table in an HTML page
 * 
 * Usage: java edu.odu.cs.gradeSync.CLI.ExtractAsHtml spreadsheetFile sheetName htmlFileName
 * 
 * @author zeil
 *
 */
public class ExtractAsHtml {


	private static Log log = LogFactory.getLog(ExtractAsHtml.class);

	private String spreadsheetFileName;
	private  String sheetName;
	private String htmlFileName;


	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		if (args.length != 3) {
			log.error("Usage: java edu.odu.cs.gradeSync.CLI.ExtractAsHtml spreadsheetFile sheetName htmlFileName");
			System.exit(1);
		}
		new ExtractAsHtml(args[0], args[1], args[2]).run();
	}

	/**
	 * Create the application.
	 * 
	 * @param spreadsheetFileName   path to an Excel spreadsheet (.xls or .xlsx)
	 * @param sheetName             name of a sheet within that workbook
	 * @param htmlFileName          path at which to store extracted HTML of that sheet
	 */
	public ExtractAsHtml(String spreadsheetFileName, String sheetName, String htmlFileName) {
		
		this.spreadsheetFileName = spreadsheetFileName;
		this.sheetName = sheetName;
		this.htmlFileName = htmlFileName;
	}

	private void run() {
		File ssFile = new File(spreadsheetFileName);
		File csvFile;
		try {
		    Spreadsheet ss = new Spreadsheet(ssFile);
		    csvFile = File.createTempFile(sheetName, ".csv", new File("."));
			csvFile.deleteOnExit();
			ss.storeCSV(csvFile, sheetName, true);
		} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
			log.error ("Unable to write sheet into temporary file", e);
			return;
		}
		
		try (CSVReader csvIn = new CSVReader(new FileReader(csvFile))) {
			try (BufferedWriter htmlOut = new BufferedWriter(new FileWriter(htmlFileName))) {
				htmlOut.write("<html><head><title>\n");
				htmlOut.write(sheetName);
				htmlOut.write("\n</title></head><body>\n");
				htmlOut.write("<table border='1'>\n<tr>\n");
				String[] headers = csvIn.readNext();
				if (headers != null) {
					for (String header: headers) {
						htmlOut.write("  <th>" + htmlEncode(header) + "</th>\n");
					}
				String[] data;
				while ((data = csvIn.readNext()) != null) {
					htmlOut.write("<tr>\n");
					for (String value: data) {
						htmlOut.write("  <td>" + htmlEncode(value) + "</td>\n");
					}
					htmlOut.write("</tr>\n");
				}
				htmlOut.write("</table>\n");
				htmlOut.write("</body>\n");
				htmlOut.write("</html>\n");
			}
		}
		} catch (IOException e) {
			log.error("I/O error", e);
		}
	}


	
	
	
	private String htmlEncode(String str) {
		StringBuffer result = new StringBuffer();
		for (int i = 0; i < str.length(); ++i) {
			char c = str.charAt(i);
			if (c == '&') {
				result.append("&amp;");
			} else if (c == '<') {
				result.append("&lt;");
			} else if (c == '>') {
				result.append("&gt;");
			} else {
				result.append(c);
			}
		}
		return result.toString();
	}




}
