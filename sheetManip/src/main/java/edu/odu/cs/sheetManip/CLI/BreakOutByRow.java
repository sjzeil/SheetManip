package edu.odu.cs.sheetManip.CLI;

import java.io.File;
import java.io.IOException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import edu.odu.cs.sheetManip.Spreadsheet;


/**
 * Utility for separating student grades recorded on a row-by-row basis
 * 
 * Usage: java edu.odu.cs.gradeSync.CLI.BreakOutByRow spreadsheetFile sheetName studentNameColumn totalsColumn
 * 
 * @author zeil
 *
 */
public class BreakOutByRow {


	private static Log log = LogFactory.getLog(BreakOutByRow.class);

	private String spreadsheetFileName;
	private  String sheetName;
	private String studentNameColumn;
	private String totalsColumn;
	private File outputDir;


	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		if (args.length != 4) {
			log.error("Usage: java edu.odu.cs.gradeSync.CLI.BreakOutByRow spreadsheetFile sheetName studentNameColumn totalsColumn");
			System.exit(1);
		}
		new BreakOutByRow(args[0], args[1], args[2], args[3], args[4]).run();
	}

	
	/**
	 * Create the application.
	 * 
	 * @param spreadsheetFileName  path to spreadsheet containing combined student data.
	 * @param sheetName            sheet in that spreadsheet to use.
	 * @param outputPath           directory where separate spreadsheets should be stored.
	 * @param studentNameColumn    column of sheet in which to find student names/identifiers .
	 * @param totalsColumn         column of sheet in which to find the student's total score for the assignment.
	 */
	public BreakOutByRow(String spreadsheetFileName, String sheetName, String outputPath, String studentNameColumn, String totalsColumn) {
		
		this.spreadsheetFileName = spreadsheetFileName;
		this.sheetName = sheetName;
		this.outputDir = new File(outputPath);
		this.studentNameColumn = studentNameColumn;
		this.totalsColumn = totalsColumn;
	}

	public void run() {
		File ssFile = new File(spreadsheetFileName);
		try {
	        Spreadsheet ss = new Spreadsheet(ssFile);
			ss.breakOutByRow(sheetName, outputDir, studentNameColumn, totalsColumn);
		} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
			log.error ("Unable to break out spreadsheet", e);
			return;
		}
	}


}
