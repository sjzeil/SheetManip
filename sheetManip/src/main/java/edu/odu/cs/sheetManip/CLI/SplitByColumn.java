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
 * Usage: java edu.odu.cs.gradeSync.CLI.SplitByColumn spreadsheetFile sheetName ouputDir studentNameRow totalsRow
 * 
 * @author zeil
 *
 */
public class SplitByColumn {


	private static Log log = LogFactory.getLog(SplitByColumn.class);

	private String spreadsheetFileName;
	private  String sheetName;
	private int studentNamesRow;
	private int totalsRow;
	private File outputDir;


	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		if (args.length != 5) {
			log.error("Usage: java edu.odu.cs.gradeSync.CLI.BreakOutByRow spreadsheetFile sheetName outputDir studentNameColumn totalsColumn");
			System.exit(1);
		}
		new SplitByColumn(args[0], args[1], args[2], 
				Integer.parseInt(args[3]), Integer.parseInt(args[4])).run();
	}

	
	/**
	 * Create the application.
	 * 
	 * @param spreadsheetFileName  path to spreadsheet containing combined student data.
	 * @param sheetName            sheet in that spreadsheet to use.
	 * @param outputPath           directory where separate spreadsheets should be stored.
	 * @param studentNameRow       row of sheet in which to find student names/identifiers .
	 * @param totalsRow            row of sheet in which to find the students'
	 *                                   total scores for the assignment.
	 */
	public SplitByColumn(String spreadsheetFileName, String sheetName, 
			String outputPath, int studentNameRow, int totalsRow) {
		
		this.spreadsheetFileName = spreadsheetFileName;
		this.sheetName = sheetName;
		this.outputDir = new File(outputPath);
		this.studentNamesRow = studentNameRow;
		this.totalsRow = totalsRow;
	}

	public void run() {
		File ssFile = new File(spreadsheetFileName);
		try {
	        Spreadsheet ss = new Spreadsheet(ssFile);
			ss.splitByColumn(sheetName, outputDir, studentNamesRow, totalsRow);
		} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
			log.error ("Unable to split spreadsheet", e);
			return;
		}
	}


}
