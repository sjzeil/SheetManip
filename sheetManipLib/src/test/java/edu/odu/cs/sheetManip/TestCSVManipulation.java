/**
 * 
 */
package edu.odu.cs.sheetManip;


import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.*;

/**
 * @author zeil
 *
 */
public class TestCSVManipulation {
    
    File testXLSFile;
    File testXLSXFile;
    File testCSVFile;

    File testXLSXFile2;
    File testCSVFile2;
    /**
     * @throws java.lang.Exception
     */
    @BeforeEach
    public void setUp() throws Exception {
        File srcDir = new File("src/test/data");
        File testDir = new File("build/testData");
        testDir.mkdirs();

        File srcSSFile = new File(srcDir, "spreadsheet1.xls");
        testXLSFile = new File(testDir, "ss.xls");
        if (testXLSFile.exists()) {
            testXLSFile.delete();
        }
        Files.copy(srcSSFile.toPath(), testXLSFile.toPath());

        srcSSFile = new File(srcDir, "spreadsheet1.xlsx");
        testXLSXFile = new File(testDir, "ss.xlsx");
        if (testXLSXFile.exists()) {
            testXLSXFile.delete();
        }
        Files.copy(srcSSFile.toPath(), testXLSXFile.toPath());

        srcSSFile = new File(srcDir, "applications.xlsx");
        testXLSXFile2 = new File(testDir, "applic.xlsx");
        if (testXLSXFile2.exists()) {
            testXLSXFile2.delete();
        }
        Files.copy(srcSSFile.toPath(), testXLSXFile2.toPath());

        File srcCsvFile = new File(srcDir, "csv1.csv");
        testCSVFile = new File(testDir, "csv.csv");
        if (testCSVFile.exists()) {
            testCSVFile.delete();
        }
        Files.copy(srcCsvFile.toPath(), testCSVFile.toPath());
        
        
        srcCsvFile = new File(srcDir, "mergeData.csv");
        testCSVFile2 = new File(testDir, "mergeData.csv");
        if (testCSVFile2.exists()) {
            testCSVFile2.delete();
        }
        Files.copy(srcCsvFile.toPath(), testCSVFile2.toPath());

}
    
    

    /**
     * @throws java.lang.Exception
     */
    @AfterEach
    public void tearDown() throws Exception {
        if (testXLSFile.exists()) {
            testXLSFile.delete();
        }
        if (testXLSXFile.exists()) {
            testXLSXFile.delete();
        }
        if (testXLSXFile2.exists()) {
            testXLSXFile2.delete();
        }
        if (testCSVFile.exists()) {
            testCSVFile.delete();
        }        
        if (testCSVFile2.exists()) {
            testCSVFile2.delete();
        }        
    }


    /**
     * Test method for {@link edu.odu.cs.sheetManip.Spreadsheet#loadCSV(java.io.File, java.lang.String)}.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    @Test
    public void testLoadCSV() throws EncryptedDocumentException, InvalidFormatException, IOException {
        Spreadsheet ss = new Spreadsheet(testXLSFile);
        ss.loadCSV(testCSVFile, "in");
        assertEquals ("C2", ss.getCellValue("in", 0, 1));
        assertEquals ("7.0", ss.getCellValue("in", 3, 3));
        assertEquals ("10.0", ss.getCellValue("out", 1, 1));        
        assertEquals ("30.0", ss.getCellValue("out", 5, 1));        
        ss.close();
    }

    /**
     * Test method for {@link edu.odu.cs.sheetManip.Spreadsheet#storeCSV(java.io.File, java.lang.String, boolean)}.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    @Test
    public void testStoreCSV() throws EncryptedDocumentException, InvalidFormatException, IOException {
        Spreadsheet ss = new Spreadsheet(testXLSXFile);
        ss.storeCSV(testCSVFile, "out", true);
        BufferedReader in = new BufferedReader (new FileReader(testCSVFile));
        String headers = in.readLine();
        assertNotNull(headers);
        assertTrue (headers.contains("Column 2."));
        assertTrue (headers.contains("Column 4."));
        String detail = in.readLine();
        assertNotNull(detail);
        assertTrue (detail.contains("Row 2."));
        assertTrue (detail.contains("3.0"));
        in.close();
        ss.close();
    }

    /**
     * Test method for {@link edu.odu.cs.sheetManip.Spreadsheet#mergeDataFromCSV(java.io.File, java.lang.String, int)}.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    @Test
    public void testMergeCSV() throws EncryptedDocumentException, InvalidFormatException, IOException {
        Spreadsheet ss = new Spreadsheet(testXLSFile);
        ss.mergeDataFromCSV(testCSVFile, "merge", 0);
        assertEquals ("R1", ss.getCellValue("merge", 1, 0));
        assertEquals ("R3", ss.getCellValue("merge", 2, 0));
        assertEquals ("R2", ss.getCellValue("merge", 3, 0));
        assertEquals ("R4", ss.getCellValue("merge", 4, 0));
        
        assertEquals ("1.0", ss.getCellValue("merge", 1, 1));
        assertEquals ("9.0", ss.getCellValue("merge", 2, 1));
        assertEquals ("9.0", ss.getCellValue("merge", 3, 1));
        assertEquals ("9.0", ss.getCellValue("merge", 4, 1));

        assertEquals ("2.0", ss.getCellValue("merge", 1, 2));
        assertEquals ("8.0", ss.getCellValue("merge", 2, 2));
        assertEquals ("8.0", ss.getCellValue("merge", 3, 2));
        assertEquals ("8.0", ss.getCellValue("merge", 4, 2));
        
        assertEquals ("3.0", ss.getCellValue("merge", 1, 3));
        assertEquals ("17.0", ss.getCellValue("merge", 2, 3));
        assertEquals ("17.0", ss.getCellValue("merge", 3, 3));
        assertEquals ("17.0", ss.getCellValue("merge", 4, 3));
        ss.close();
        
    }

    /**
     * Test method for {@link edu.odu.cs.sheetManip.Spreadsheet#mergeDataFromCSV(java.io.File, java.lang.String, int)}.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    @Test
    public void testMergeCSV2() throws EncryptedDocumentException, InvalidFormatException, IOException {
        Spreadsheet ss = new Spreadsheet(testXLSXFile2);
        ss.mergeDataFromCSV(testCSVFile2, "SPEAK-GTAI", 3);
        assertEquals ("szeil", ss.getCellValue("SPEAK-GTAI", 1, 3));
        assertEquals ("jdoe001", ss.getCellValue("SPEAK-GTAI", 2, 3));
        assertEquals ("jsmit999", ss.getCellValue("SPEAK-GTAI", 3, 3));
        assertNull (ss.getCellValue("SPEAK-GTAI", 4, 3));
        
        ss.close();
        
    }
    
    /**
     * Test method for {@link edu.odu.cs.sheetManip.Spreadsheet#breakOutByRow(java.lang.String, java.io.File, java.lang.String, java.lang.String)}.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    @Test
    public void testBreakOutByRow() throws EncryptedDocumentException, InvalidFormatException, IOException {
        Spreadsheet ss = new Spreadsheet(testXLSXFile);
        File testDir = new File("build/testData");
        ss.breakOutByRow("in", testDir, "A", "D");
        File ss1 = new File(testDir, "Row 2.xls");
        File ss2 = new File(testDir, "Row 3.xls");
        File ss3 = new File(testDir, "Row 4.xls");
        assertTrue(ss1.exists());
        
        ss.close();
        ss = new Spreadsheet(ss1);
        assertEquals ("Column 2", ss.getCellValue("in", 0, 1));
        assertEquals ("Row 2", ss.getCellValue("in", 1, 0));
        assertEquals ("0.0", ss.getCellValue("in", 1, 1));
        assertEquals ("0.0", ss.getCellValue("in", 1, 3));
        
        ss1.delete();
        assertTrue(ss2.exists());
        ss2.delete();
        assertTrue(ss3.exists());
        ss3.delete();
        
        ss.close();
    }

}
