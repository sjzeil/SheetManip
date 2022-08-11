package edu.odu.cs.sheetManip.CLI;

import java.io.File;
import java.io.IOException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import edu.odu.cs.sheetManip.Spreadsheet;


/**
 * Utility for merging a sheet with data from a CSV file.
 * 
 * Usage: java edu.odu.cs.gradeSync.CLI.ExtractAsCsv spreadsheetFile sheetName csvFileName
 * 
 * @author zeil
 *
 */
public class MergeFromCsv {


	private static Log log = LogFactory.getLog(MergeFromCsv.class);

	private String spreadsheetFileName;
	private  String sheetName;
	private String csvFileName;
	private int keyColumn;


	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		if (args.length != 4) {
			log.error("Usage: java edu.odu.cs.gradeSync.CLI.LoadFromCsv spreadsheetFile sheetName csvFileName keyColumn");
			System.exit(1);
		}
		int keyColumn;
		if (Character.isAlphabetic(args[3].charAt(0))) {
			keyColumn = args[3].charAt(0) - 'A';
		} else {
			keyColumn = Integer.parseInt(args[3]);
		}
		new MergeFromCsv(args[0], args[1], args[2], keyColumn).run();
	}

	/**
	 * Create the application.
	 * 
	 * @param spreadsheetFileName   path to an Excel spreadsheet (.xls or .xlsx)
	 * @param sheetName             name of a sheet within that workbook to be overwritten
	 * @param csvFileName           path to a CSV file to insert into that sheet
	 * @param keyCol column of spreadsheet (A=0, B-=1, ...) containing the merge key
	 */
	public MergeFromCsv(String spreadsheetFileName, String sheetName, String csvFileName, int keyCol) {
		
		this.spreadsheetFileName = spreadsheetFileName;
		this.sheetName = sheetName;
		this.csvFileName = csvFileName;
		this.keyColumn = keyCol;
	}

	private void run() {
		File ssFile = new File(spreadsheetFileName);
		File csvFile = new File(csvFileName);
		try {
		    Spreadsheet ss = new Spreadsheet(ssFile);
		    ss.mergeDataFromCSV(csvFile, sheetName, keyColumn);
		} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
			log.error ("Unable to load sheet from temporary file", e);
			return;
		}
	}


}
