package edu.odu.cs.sheetManip;
/**
 * 
 */


import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;



/**
 * A grading spreadsheet/workbook from which individual sheets can be extracted
 * as CSV or have date loaded from CSV.
 * 
 *  
 * @author zeil
 *
 */
public class Spreadsheet {

    /**
     * Log for error messages
     */
    private static Log log = LogFactory.getLog(Spreadsheet.class);

    /**
     * Excel file being manipulated.
     */
    private File xlsFile;
    
    /**
     * Stream used to load an existing spreadsheet.
     */
    private FileInputStream inStream;
    
    /**
     * Workbook contained in that file.
     */
    private Workbook wb;


    /**
     * Create a grade spreadsheet object mapping onto an existing
     * Excel spreadsheet file (xls or xslx format).
     * 
     * @param excelFile location of an Excel spreadsheet. Creates a
     *    new spreadsheet if no such file exists.
     *    
     * @throws IOException if file cannot be read
     * @throws InvalidFormatException  if file is a not a spreadsheet
     * @throws EncryptedDocumentException  if file is encrypted?
     */
    public Spreadsheet(final File excelFile) throws EncryptedDocumentException, InvalidFormatException, IOException {
        this.xlsFile = excelFile;
        inStream = new FileInputStream(excelFile);
        wb = WorkbookFactory.create(inStream);
    }


    /**
     * Release internal resources held by this spreadsheet.
     * 
     * @throws IOException if unable to close
     */
    public void close() throws IOException {
        if (inStream != null)
            inStream.close();
        wb.close();
    }
    
    /**
     * Retrieve the value of a spreadsheet cell, expressed as a string.
     * 
     * @param sheetName name of the sheet from which to retrieve
     * @param rowNum index of the row from which to retrieve
     * @param column index of the column from which to retrieve
     * @return formula string in that cell, or null if cell does not exist
     */
    public String getCellValue(String sheetName, int rowNum, int column) {
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            return null;
        }
        Row row = sheet.getRow(rowNum);
        if (row == null) {
            return null;
        }
        Cell c = row.getCell(column);
        if (c == null) {
            return null;
        }
        return evaluateCell(c, wb);
    }



    /**
     * Removes all data entries from a sheet.
     *
     * @param workbook  which workbook to employ
     * @param sheetName name of sheet within workbook
     */
    private void clearSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet != null) {
            // Decide which rows to process
            int rowStart = sheet.getFirstRowNum();
            int rowEnd = sheet.getLastRowNum();

            for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
                Row r = sheet.getRow(rowNum);
                if (r != null) {
                    int lastColumn = r.getLastCellNum();

                    for (int cn = 0; cn < lastColumn; cn++) {
                        Cell c = r.getCell(cn, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (c != null) {
                            c.setBlank();
                        }
                    }
                }
            }
        }
    }

    /**
     * Write a workbook into a file.
     * 
     * @param workbook spreadsheet to be written
     */
    private void saveWorkBook(Workbook workbook) {
        workbook.setForceFormulaRecalculation(true);
        if (inStream != null)
            try {
                inStream.close();
            } catch (IOException e1) {
                log.warn("Unexpected error closing input stream", e1);
            }
        if (xlsFile.exists()) {
            String xlsFileAbs = /* xlsFile.getAbsolutePath(); */ xlsFile.getPath();
            File backup = new File(xlsFileAbs + ".bak");
            if (!backup.exists()) {
                try {
                    Files.move(xlsFile.toPath(), backup.toPath());
                } catch (IOException e) {
                    log.error("Unable to save old file as a backup", e);
                }
            } else {
                xlsFile.delete();
            }
        }
        FileOutputStream fileOut = null;
        try  {
            fileOut  = new FileOutputStream(xlsFile);
            workbook.write(fileOut);
        } catch (IOException e) {
            log.error("Problem writing out spreadsheet", e);
        } finally {
            if (fileOut != null) {
                try {
                    fileOut.close();
                } catch (IOException e) {
                    log.warn("Problem closing spreadsheet", e);
                }
            }
        }

    }



    /**
     * Replace the contents of a sheet by the contents of a CSV file,
     * updating calculations elsewhere in the spreadsheet afterwards.
     *  
     * @param csvFile  a CSV file with one sheet's worth of spreadsheet data.
     * @param sheetName the name of the sheet into which to place that data.
     *    Normally this sheet should already exist (other wise there could
     *    be no calculations upon it, 
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    public void loadCSV(final File csvFile, final String sheetName) 
            throws EncryptedDocumentException, InvalidFormatException, IOException {
        clearSheet(wb, sheetName);
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            sheet = wb.createSheet(sheetName);
        }

        BufferedReader rdr = new BufferedReader(new FileReader(csvFile));
        CSVReader csvIn = new CSVReader(rdr);
        String[] line = csvIn.readNext();
        int rowNum = 0;
        while (line != null) {
            Row r = sheet.getRow(rowNum);
            if (r == null) {
                r = sheet.createRow(rowNum);
            }
            for (int colNum = 0; colNum < line.length; ++colNum) {
                String value = line[colNum];
                Cell c = r.getCell(colNum, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (c == null) {
                    c = r.createCell(colNum);
                }
                try {
                    Double d = Double.parseDouble(value);
                    //c.setCellType(CellType.NUMERIC);
                    c.setCellValue(d);
                } catch (NumberFormatException ex) {
                    //c.setCellType(CellType.STRING);
                    c.setCellValue(value);
                }


            }
            ++rowNum;
            line = csvIn.readNext();
        }
        csvIn.close();
        saveWorkBook(wb);
    }



    /**
     * Merge the contents of a sheet with the contents of a CSV file,
     * updating calculations elsewhere in the spreadsheet afterwards.
     * 
     * The merge observes the following rules:
     * 
     * 1. Cells that contain formulae are left unchanged by the merge process.
     *    Only cells containing simple data values are replaced.
     * 2. The row in the sheet to be replaced by a line from the CSV file is
     *    determined by comparing the values in the yey column. If the CSV key
     *    matches a row in the sheet, then row is replaced by the CSV data.
     *    If a CSV line has a key value matching no row of the sheet, then
     *    the first sheet row with an empty key field is used.   
     *  
     * @param csvFile  a CSV file with one sheet's worth of spreadsheet data.
     * @param sheetName the name of the sheet into which to place that data.
     *    Normally this sheet should already exist (other wise there could
     *    be no calculations upon it,
     * @param int keyColumn  which column (A=1, B=2, ...) holds the key used
     *         to match CSV rows to the corresponding spreadsheet row 
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    public void mergeDataFromCSV(final File csvFile, 
            final String sheetName, 
            int keyColumn) 
                    throws EncryptedDocumentException, InvalidFormatException, IOException {
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            sheet = wb.createSheet(sheetName);
        }

        BufferedReader rdr = new BufferedReader(new FileReader(csvFile));
        CSVReader csvIn = new CSVReader(rdr);
        String[] line = csvIn.readNext();
        while (line != null) {
            String keyValue = line[keyColumn];
            if (!keyValue.equals("")) {
                int rowNum = findMatchingRow(sheet, keyValue, keyColumn);

                Row r = sheet.getRow(rowNum);
                if (r == null) {
                    r = sheet.createRow(rowNum);
                }

                for (int colNum = 0; colNum < line.length; ++colNum) {
                    String value = line[colNum];
                    Cell c = r.getCell(colNum, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if (c == null) {
                        c = r.createCell(colNum);
                    }
                    CellType ctype = c.getCellType();
                    if (ctype != CellType.FORMULA) {                
                        try {
                            Double d = Double.parseDouble(value);
                            //c.setCellType(CellType.NUMERIC);
                            c.setCellValue(d);
                        } catch (NumberFormatException ex) {
                            //c.setCellType(CellType.STRING);
                            c.setCellValue(value);
                        }
                    }
                }
            }
            line = csvIn.readNext();
        }
        csvIn.close();
        saveWorkBook(wb);
    }


    /**
     * Scan a sheet for a row containing a value equal to keyValue in column
     * keyColumn.
     * 
     * @param sheet   a sheet from a workbook
     * @param keyValue  a string value
     * @param keyColumn column in which keyValue may be located
     * @return  Row number that matches the given key. If no matching
     *          row is found, returns 1+k where k is the last row number
     *          containing a non-empty value in the key column. 
     */
    private int findMatchingRow(Sheet sheet, String keyValue, int keyColumn) {
        int rowEnd = sheet.getLastRowNum();
        int lastFilledRow = -1; 
        for (int rowNum = 0; rowNum <= rowEnd; ++rowNum) {
            Row ssRow = sheet.getRow(rowNum);
            int lastCol = (ssRow == null) ? 0 : Math.max(ssRow.getLastCellNum(), 0);
            if (lastCol < keyColumn) continue;

            Cell c = ssRow.getCell(keyColumn);
            if (c == null) continue;
            
            CellType cellType = c.getCellType();
            if (cellType == CellType.STRING) {
                lastFilledRow = rowNum;
                if (keyValue.equals(c.getStringCellValue())) {
                    return rowNum;
                }
            } else if (cellType == CellType.NUMERIC) {
                lastFilledRow = rowNum;
                String cellValue = String.format("%d", Math.round(c.getNumericCellValue()));
                if (keyValue.equals(cellValue)) {
                    return rowNum;
                }
            } else if (cellType == CellType.FORMULA
                    || cellType == CellType.BOOLEAN) {
                lastFilledRow = rowNum;
            }
        }
        return lastFilledRow+1;
    }



    /**
     * Evaluate the formula in a spreadsheet cell.
     *
     * @param c the cell
     * @param workbook the spreadsheet
     * @return string representation of the cell's value
     */
    private String evaluateCell(Cell c, Workbook workbook) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        String value = "";
        if (c != null) {
            CellValue cellValue = null;
            try {
                cellValue = evaluator.evaluate(c);
            } catch (Exception ex) {
                return "**err**";
            }
            if (cellValue != null) {
                CellType cellType = cellValue.getCellType();
                switch (cellType) {
                case STRING:
                    value = cellValue.getStringValue();
                    break;
                case NUMERIC:
                    value = "" + cellValue.getNumberValue();
                    if (value.matches("^[+-]?[0-9]*[.][0-9][0-9][0-9][0-9]*")) {
                        Double d = Double.parseDouble(value);
                        value = String.format("%.2f", d);
                    }
                    break;
                case BOOLEAN:
                    value = "" + cellValue.getBooleanValue();
                    break;
                default:
                    return "??";
                }
            }
        }
        return value;
    }



    /**
     * Write the contents (values only) of a sheet by into a CSV file.
     *  
     * @param csvFile  location of the desired CSV file.
     * @param sheetName the name of the sheet from which to obtain the data.
     * @param  skipInvalidDataRows if true, omit any row containing a non-empty
     *                             value that is not a valid number or string.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    public void storeCSV(final File csvFile, 
            final String sheetName,
            boolean skipInvalidDataRows) 
                    throws EncryptedDocumentException, InvalidFormatException, IOException {

        List<String[]> csvContents = evaluateSheet(
                sheetName, 
                skipInvalidDataRows);

        BufferedWriter wtr = new BufferedWriter(new FileWriter(csvFile));
        CSVWriter csvOut = new CSVWriter(wtr);
        csvOut.writeAll(csvContents, false);
        csvOut.close();
    }

    
    
    /**
     * Copy non-empty, non-formula cells from a sheet of another spreadsheet
     * into this one.
     * 
     * @param from spreadsheet to copy from
     * @param fromSheetName which sheet in that spreadsheet to copy from
     * @param intoSheetName which sheet in this spreadsheet to copy into
     */
    public void copySheetData(Spreadsheet from, String fromSheetName, String intoSheetName)
    {
        Sheet intoSheet = wb.getSheet(intoSheetName);
        Sheet fromSheet = from.wb.getSheet(fromSheetName);

        int rowEnd = fromSheet.getLastRowNum();
        for (int rowNum = 0; rowNum <= rowEnd; ++rowNum) {
            Row ssRow = fromSheet.getRow(rowNum);
            if (ssRow == null) continue; 
            Row intoRow = intoSheet.getRow(rowNum);
            if (intoRow == null) {
                intoRow = intoSheet.createRow(rowNum);
            }
            int lastCol = (ssRow == null) ? 0 : Math.max(ssRow.getLastCellNum(), 0);
            for (int colNum = 0; colNum < lastCol; ++colNum) {
                Cell c = ssRow.getCell(colNum);
                if (c == null) continue;
                CellType ctype = c.getCellType();
                switch (ctype) {
                case FORMULA: break;
                case ERROR: break;
                case BLANK:
                case STRING: 
                {
                    String v = c.getStringCellValue();
                    Cell c2 = intoRow.getCell(colNum, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if (c2 == null) {
                        c2 = intoRow.createCell(colNum);
                    }
                    c2.setCellValue(v);
                }
                break;
                case NUMERIC:
                {
                    double d = c.getNumericCellValue();
                    Cell c2 = intoRow.getCell(colNum, 
                            Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if (c2 == null) {
                        c2 = intoRow.createCell(colNum);
                    }
                    c2.setCellValue(d);
                }
                default:
                    break;
                }
            }
        }
        saveWorkBook(wb);
    }

        

    
    
    /**
     * Fetch the contents (values only) of a sheet as a list (by rows) of
     * arrays (by column) of strings.
     *  
     * @param sheetName the name of the sheet from which to obtain the data.
     * @param  skipInvalidDataRows if true, omit any row containing a non-empty
     *                             value that is not a valid number or string.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     * 
     * @return list of rows
     */
    public List<String[]> evaluateSheet(
            final String sheetName,
            boolean skipInvalidDataRows) 
                    throws EncryptedDocumentException, InvalidFormatException, IOException {
        Sheet sheet = wb.getSheet(sheetName);

        int rowEnd = sheet.getLastRowNum();
        List<String[]> csvContents = new ArrayList<String[]>();
        for (int rowNum = 0; rowNum <= rowEnd; ++rowNum) {
            Row ssRow = sheet.getRow(rowNum);
            int lastCol = (ssRow == null) ? 0 : Math.max(ssRow.getLastCellNum(), 0);
            String[] row = new String[lastCol];
            boolean rowIsValid = true;
            boolean rowIsEmpty = true;
            for (int colNum = 0; colNum < lastCol; ++colNum) {
                Cell c = ssRow.getCell(colNum);
                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                String value = "";
                if (c != null) {
                    CellValue cellValue = null;
                    try {
                        cellValue = evaluator.evaluate(c);
                    } catch (Exception ex) {
                        rowIsValid = false;
                        colNum = lastCol;
                    }
                    if (cellValue != null) {
                        CellType cellType = cellValue.getCellType();
                        switch (cellType) {
                        case STRING:
                            value = cellValue.getStringValue();
                            break;
                        case NUMERIC:
                            value = "" + cellValue.getNumberValue();
                            if (value.matches("^[+-]?[0-9]*[.][0-9][0-9][0-9][0-9]*")) {
                                Double d = Double.parseDouble(value);
                                value = String.format("%.2f", d);
                            }
                            break;
                        case BOOLEAN:
                            value = "" + cellValue.getBooleanValue();
                            break;
                        default:
                            if (skipInvalidDataRows) {
                                rowIsValid = false;
                                colNum = lastCol;
                            }
                        }
                    }
                }
                if (c != null && rowIsValid) {
                    row[colNum] = value;
                    rowIsEmpty = false;
                }
            }
            if (rowIsValid && !rowIsEmpty) {
                csvContents.add(row);
            }
        }

        return csvContents;
    }


    /**
     * Scans the given spreadsheet. The column identified as studentNameColumn is presumed to contain
     * names/identifiers for distinct students.  For each such name N encountered, a new spreadsheet 
     * outputDir/N.xls is created containing 1) all rows of this sheet that have an empty cell in that
     * identifier column, 2) all rows of this sheet having N in that column, and 3) a new row with "Total:"
     * in column A and the last entry from the totalColumn in a row having N in the identifier column.
     *    
     * @param sheetName   which sheet of this spreadsheet to process
     * @param outputDir  directory to contain the newly generated spreadsheets
     * @param studentNameColumn  column containing student identifiers
     * @param totalsColumn    column containing entry for the Total: row
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    public void breakOutByRow(String sheetName, File outputDir, String studentNameColumn, String totalsColumn) 
            throws EncryptedDocumentException, InvalidFormatException, IOException {
        Sheet sheet = wb.getSheet(sheetName);

        int studentNameColNumber = CellReference.convertColStringToIndex(studentNameColumn);
        int totalsColNumber = CellReference.convertColStringToIndex(totalsColumn);

        int rowEnd = sheet.getLastRowNum();
        Set<String> studentNames = new TreeSet<String>();
        for (int rowNum = 1; rowNum <= rowEnd; ++rowNum) {
            Row ssRow = sheet.getRow(rowNum);
            int lastCol = (ssRow == null) ? -1 : ssRow.getLastCellNum();
            if (lastCol >= studentNameColNumber) {
                Cell c = ssRow.getCell(studentNameColNumber);
                String name = evaluateCell(c, wb).trim();
                if (name.length() > 0) {
                    studentNames.add(name);
                }
            }
        }

        //System.err.println("Names: " + studentNames);

        for (String studentName: studentNames) {
            File studentSSFile = new File(outputDir, studentName + ".xls");
            Workbook studentWB = new HSSFWorkbook();
            Sheet gradesSheet = studentWB.createSheet(sheetName);

            String totalValue = "";
            int nRows = 0;

            for (int rowNum = 0; rowNum <= rowEnd; ++rowNum) {
                Row ssRow = sheet.getRow(rowNum);
                String idValue = "";
                int lastCol = (ssRow == null) ? -1 : ssRow.getLastCellNum();
                if (lastCol >= studentNameColNumber) {
                    Cell c = ssRow.getCell(studentNameColNumber);
                    idValue = evaluateCell(c, wb).trim();
                }
                if (idValue.length() == 0 || idValue.equals(studentName)) {
                    // Copy this row into the student spreadsheet
                    Row studentRow = gradesSheet.createRow(nRows);
                    for (int col = 0; col <= lastCol; ++col) {
                        Cell c = ssRow.getCell(col);
                        if (c != null) {
                            String value = evaluateCell(c, wb);
                            Cell cNew = studentRow.createCell(col, CellType.STRING);
                            cNew.setCellValue(value);
                        }
                    }

                    if (lastCol >= totalsColNumber) {
                        Cell c = ssRow.getCell(totalsColNumber);
                        totalValue = evaluateCell(c, wb).trim();
                    }   
                    ++nRows;
                }
            }    
            Row totalsRow = gradesSheet.createRow(nRows + 1);
            totalsRow.createCell(0, CellType.STRING).setCellValue("Total:");
            totalsRow.createCell(1, CellType.STRING).setCellValue(totalValue);


            FileOutputStream fileOut = new FileOutputStream(studentSSFile);
            studentWB.write(fileOut);
            fileOut.close();
            studentWB.close();
        }

    }


    /**
     * This function is used for grade spreadsheets that contain multiple
     * students in a column by column format (e.g., a team project)
     * from which we need to obtain separate spreadsheets for each student.
     * 
     * For each student, a copy of the spreadsheet is generated that is
     * identical to the original spreadsheet except that, in a row containing
     * student names/identifiers, all other names are removed (column A is
     * exempt because it typically contains a generic label) and, in a
     * row containing the students' total scores, all other scores are removed.
     *  
     * @param sheetName   which sheet of this spreadsheet to process
     * @param outputDir  directory to contain the newly generated spreadsheets
     * @param studentNamesRow integer identifier of the row containing student
     *                        names/identifiers. Each such ID is used as the
     *                        name of the newly generated spreadsheet.
     * @param totalsRowNum integer identifier of the row containing student scores.
     * @throws IOException if spreadsheet cannot be closed and copied
     * @throws InvalidFormatException if copied spreadsheet cannot be opened
     * @throws EncryptedDocumentException if copied spreadsheet cannot be opened
     */
	public void splitByColumn(String sheetName, File outputDir,
			int studentNamesRow, int totalsRowNum) throws IOException, EncryptedDocumentException, InvalidFormatException {
        Sheet sheet = wb.getSheet(sheetName);
		String extension = xlsFile.getName().substring(xlsFile.getName().lastIndexOf('.'));

        int studentNameRowNumber = studentNamesRow - 1;
        int totalsRowNumber = totalsRowNum - 1;

        Row namesRow = sheet.getRow(studentNameRowNumber);
        Row totalsRow = sheet.getRow(totalsRowNumber);
        
        int colEnd = Math.max(namesRow.getLastCellNum(), totalsRow.getLastCellNum());
        Map<String, Integer> studentNames = new TreeMap<>();
        for (int columnNum = 1; columnNum < colEnd; ++columnNum) {
            Cell nameCell = namesRow.getCell(columnNum);
            Cell scoreCell = totalsRow.getCell(columnNum);
            String name = evaluateCell(nameCell, wb).trim();
            String score = evaluateCell(scoreCell, wb).trim();
            
            try {
            	Double.parseDouble(score);
            } catch (NumberFormatException ex) {
            	name = "";
            }
            if (!name.contentEquals("")) /* and score is numeric */ {
            	studentNames.put(name, columnNum);
            }
        }
        wb.close();
        //System.err.println("Names: " + studentNames);

        for (String studentName: studentNames.keySet()) {
            File studentSSFile = new File(outputDir, studentName + extension);
            File studentSSTemp = new File(outputDir, studentName + "_temp" + extension);
            Files.copy(xlsFile.toPath(), studentSSTemp.toPath());
            
            Workbook studentWB = WorkbookFactory.create(studentSSTemp);
            sheet = studentWB.getSheet(sheetName);
            namesRow = sheet.getRow(studentNameRowNumber);
            totalsRow = sheet.getRow(totalsRowNumber);
            
            int preservedColumn = studentNames.get(studentName);
            for (int i = 1; i < colEnd; ++i) {
            	if (i != preservedColumn) {
            		Cell nameCell = namesRow.getCell(i);
                    Cell scoreCell = totalsRow.getCell(i);
                    if (nameCell != null)
                        nameCell.setBlank();
                    if (scoreCell != null)
                        scoreCell.setBlank();
            	}
            }
            FileOutputStream fileOut = new FileOutputStream(studentSSFile);
            studentWB.write(fileOut);
            fileOut.close();
            try {
                studentWB.close();
            } catch (IOException e) {
                log.error("Possible problem closing spreadsheet "
                    + studentSSFile + ": " + e.getMessage());
            }
            studentSSTemp.delete();
        }

	}


	/**
	 * Scans the spreadsheet for cells containing strings of the form
	 * leftDelimiter + pName + rightDelimiter, where name pName is the name
	 * of a property in properties.   Replaces the value of any such cell by
	 * the value of that property. If the value can be parsed as a number, the
	 * cell is set to numeric. otherwise the property value is inserted as a
	 * string.
	 * 
	 * @param properties  a collection of named properties
	 * @param leftDelimiter string to expect to the left of a property name
	 * @param rightDelimiter string to expect to the right of a property name
	 */
	public void injectProperties(Properties properties,
	        String leftDelimiter, String rightDelimiter) {

	    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
	        Sheet sheet = wb.getSheetAt(i);
	        int rowEnd = sheet.getLastRowNum();
	        for (int rowNum = 0; rowNum <= rowEnd; ++rowNum) {
	            Row ssRow = sheet.getRow(rowNum);
	            int lastCol = (ssRow == null) ? 0 : Math.max(ssRow.getLastCellNum(), 0);
	            for (int colNum = 0; colNum < lastCol; ++colNum) {
	                Cell cell = ssRow.getCell(colNum);
	                if (cell == null) continue;
	                if (cell.getCellType() != CellType.STRING) continue;
	                String pName = cell.getStringCellValue();
	                if (pName.startsWith(leftDelimiter) 
	                        && pName.endsWith(rightDelimiter)) {
	                    pName = pName.substring(leftDelimiter.length());
	                    pName = pName.substring(0,
	                            pName.length() - rightDelimiter.length());
	                    String value = properties.getProperty(pName);
	                    if (value != null) {
	                        try {
	                            double d = Double.parseDouble(value);
	                            cell.setCellValue(d);
	                        } catch (NumberFormatException ex) {
	                            // Non-numeric value - insert as string literal
	                            cell.setCellValue(value);
	                        }
	                    }
	                }
	            }
	        }
	    }
        saveWorkBook(wb);
	}

	
	/**
	 * Clears a rectangular region of the specified sheet.
	 * 
	 * @param sheetName sheet in which to clear
	 * @param ulRow  row of upper left corner of region to clear
	 * @param ulCol  column of upper left corner of region to clear
	 * @param lrRow  row of lower right corner of region to clear
	 * @param lrCol  column of lower right corner of region to clear
	 */
    public void clear(String sheetName, 
            int ulRow, int ulCol,
            int lrRow, int lrCol) {
        Sheet sheet = wb.getSheet(sheetName);

        for (int rowNum = ulRow; rowNum <= lrRow; ++rowNum) {
            Row ssRow = sheet.getRow(rowNum);
            if (ssRow == null) continue; 
            for (int colNum = ulCol; colNum <= lrCol; ++colNum) {
                Cell c = ssRow.getCell(colNum);
                if (c == null) continue;
                ssRow.removeCell(c);
            }
        }
        saveWorkBook(wb);
    }


    /**
     * @return a list of the names of all sheets in the spreadsheet file.
     */
    public List<String> getSheetNames() {
        List<String> results = new ArrayList<>();
        for (Sheet sheet: wb) {
            String name = sheet.getSheetName();
            results.add(name);
        }
        return results;
    }


    /**
     * Renders a sheet as an HTML table.
     *  
     * @param sheetName the name of the sheet from which to obtain the data.
     * @param  skipInvalidDataRows if true, omit any row containing a non-empty
     *                             value that is not a valid number or string.
     * @param row1Prefix string to place in front of each value in row 1,
     *      e.g., "<b>"
     * @param row1Suffix string to place after each value in row 1,
     *      e.g., "</b>"
     * @param colAPrefix string to place in front of each value in column A,
     *      in rows > 1, e.g., "<i>"
     * @param colASuffix string to place after each value in column A,
     *      in rows > 1, e.g., "</i>"
     * @return the text of an HTML table containing the values from the
     *          selected sheet.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    public String sheetToHTML(
            final String sheetName,
            boolean skipInvalidDataRows,
            String row1Prefix,
            String row1Suffix,
            String colAPrefix,
            String colASuffix
            ) 
        throws EncryptedDocumentException, InvalidFormatException, IOException {

        List<String[]> contents = evaluateSheet(
                sheetName, 
                skipInvalidDataRows);

        StringBuilder out = new StringBuilder();
        out.append("<table border='1'>\n");
        int rowNum = 0;
        for (String[] row: contents) {
            out.append("<tr>");
            for (int col = 0; col < row.length; ++col) {
                String prefix = (rowNum == 0) ? row1Prefix 
                    : ((col == 0) ? colAPrefix : "");
                out.append("<td>");
                out.append(prefix);
                String suffix = (rowNum == 0) ? row1Suffix 
                    : ((col == 0) ? colASuffix : "");
                String value = row[col];
                if (value == null) {
                    value = "";
                }
                boolean isANumber = false;
                try {
                    Double.parseDouble(value);
                    isANumber = true;
                } catch (NumberFormatException ex) {
                    isANumber = false;
                }
                if (isANumber && value.endsWith(".0")) {
                    value = value.substring(0, value.length()-2);
                }
                if (value.contains("\n") || value.contains("\r")) {
                    value = "\n<pre>" + htmlSafe(value) + "</pre>\n";
                } else {
                    value = htmlSafe(value);
                }
                out.append(value);
                out.append(suffix);
                out.append("</td>");
            }
            out.append("</tr>\n");
            ++rowNum;
        }
        out.append("</table>\n");
        return out.toString();
    }


    /**
     * Provides an HTML summary of the entire spreadsheet.
     * @param title title for the page.
     * @param  skipInvalidDataRows if true, omit any row containing a non-empty
     *                             value that is not a valid number or string.
     * @param row1Prefix string to place in front of each value in row 1,
     *      e.g., "<b>"
     * @param row1Suffix string to place after each value in row 1,
     *      e.g., "</b>"
     * @param colAPrefix string to place in front of each value in column A,
     *      in rows > 1, e.g., "<i>"
     * @param colASuffix string to place after each value in column A,
     *      in rows > 1, e.g., "</i>"
     * @return the text of an HTML table containing the values from the
     *          selected sheet.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     * @return a complete HTML page containing a table for each sheet.
     */
    public String toHTML(
            final String title,
            boolean skipInvalidDataRows,
            final String row1Prefix,
            final String row1Suffix,
            final String colAPrefix,
            final String colASuffix
            ) 
        throws EncryptedDocumentException, InvalidFormatException, IOException {

        StringBuilder out = new StringBuilder();
        out.append ("<html>\n<head>\n<title>");
        out.append (htmlSafe(title));
        out.append ("</title>\n</head>\n<body><h1>");
        out.append (htmlSafe(title));
        out.append("</h1>\n");

        List<String> sheetNames = getSheetNames();

        for (String sheetName: sheetNames) {
            String table = sheetToHTML(sheetName, skipInvalidDataRows, 
                row1Prefix, row1Suffix, colAPrefix, colASuffix);
            out.append("<h2>" + htmlSafe(sheetName) + "</h2>\n");
            out.append(table);
        }
        out.append("\n</body>\n</html>\n");
        return out.toString();
}

    /**
     * Make a string "HTML safe" by encoding &, <, and > characters.
     * @param str a string intended for injection into an HTML page.
     * @return encoded version of str
     */
    private String htmlSafe(String title) {
        String result = title.replace("&", "&amp;");
        result = result.replace("<", "&lt;");
        result = result.replace(">", "&gt;");
        return result;
    }
}
