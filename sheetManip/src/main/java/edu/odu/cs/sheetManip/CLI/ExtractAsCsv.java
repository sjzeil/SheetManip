package edu.odu.cs.sheetManip.CLI;

import java.io.File;
import java.io.IOException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import edu.odu.cs.sheetManip.Spreadsheet;


/**
 * Utility for extracting a sheet as a CSV file.
 * 
 * Usage: java edu.odu.cs.gradeSync.CLI.ExtractAsCsv spreadsheetFile sheetName csvFileName
 * 
 * @author zeil
 *
 */
public class ExtractAsCsv {


	private static Log log = LogFactory.getLog(ExtractAsCsv.class);

	private String spreadsheetFileName;
	private  String sheetName;
	private String csvFileName;


	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		if (args.length != 3) {
			log.error("Usage: java edu.odu.cs.gradeSync.CLI.ExtractAsCsv spreadsheetFile sheetName csvFileName");
			System.exit(1);
		}
		new ExtractAsCsv(args[0], args[1], args[2]).run();
	}

	/**
	 * Create the application.
	 * 
	 * @param spreadsheetFileName   path to an Excel spreadsheet (.xls or .xlsx)
	 * @param sheetName             name of a sheet within that workbook
	 * @param csvFileName           path at which to store extracted CSV of that sheet
	 */
	public ExtractAsCsv(String spreadsheetFileName, String sheetName, String csvFileName) {
		
		this.spreadsheetFileName = spreadsheetFileName;
		this.sheetName = sheetName;
		this.csvFileName = csvFileName;
	}

	public void run() {
		File ssFile = new File(spreadsheetFileName);
		File csvFile = new File(csvFileName);
		try {
		    Spreadsheet ss = new Spreadsheet(ssFile);
		    ss.storeCSV(csvFile, sheetName, true);
		} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
			log.error ("Unable to write sheet into temporary file", e);
			return;
		}
	}


}
