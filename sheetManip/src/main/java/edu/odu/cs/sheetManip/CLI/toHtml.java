package edu.odu.cs.sheetManip.CLI;

import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintStream;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.opencsv.CSVReader;

import edu.odu.cs.sheetManip.Spreadsheet;


/**
 * Utility for summarizing the entire spreadsheet in an HTML file, one table
 * per sheet.
 * 
 * Usage: java edu.odu.cs.gradeSync.CLI.toHtml spreadsheetFile [title]
 * 
 * @author zeil
 *
 */
public class toHtml {


	private static Log log = LogFactory.getLog(toHtml.class);

	private String spreadsheetFileName;
	private  String title;
	

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		if (args.length != 2 && args.length != 1) {
			log.error("Usage: java edu.odu.cs.gradeSync.CLI.toHtml spreadsheetFile [title]");
			System.exit(1);
		}
        if (args.length == 1) {
            new toHtml(args[0], null).run();
        } else {
		    new toHtml(args[0], args[1]).run();
        }
	}

	/**
	 * Create the application.
	 * 
	 * @param spreadsheetFileName   path to an Excel spreadsheet (.xls or .xlsx)
	 * @param title             name of a sheet within that workbook
	 */
	public toHtml(String spreadsheetFileName, String theTitle) {
		
		this.spreadsheetFileName = spreadsheetFileName;
		this.title = theTitle;
	}

	public void run() {
        // Apache POI issues an annoying warning directly to System.out
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream strOut = new PrintStream(baos);
        PrintStream oldOut = System.out;

        File ssFile = new File(spreadsheetFileName);

        System.setOut(strOut);


        if (title == null) {
            title = ssFile.getName();
        }
		try {
            Spreadsheet ss = new Spreadsheet(ssFile);
            String htmlPage = ss.toHTML(title, true,
                "<b>", "</b>", "<i>", "</i>");
                ss.close();
                System.setOut(oldOut);
            System.out.println(htmlPage);
		} catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
			log.error ("Unable to read spreadsheet", e);
			return;
		}
	}






}
