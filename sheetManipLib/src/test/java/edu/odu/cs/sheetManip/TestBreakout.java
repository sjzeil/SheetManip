/**
 * 
 */
package edu.odu.cs.sheetManip;

import java.io.File;
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
public class TestBreakout {
    
    File testXLSXFile;

    /**
     * @throws java.lang.Exception
     */
    @BeforeEach
    public void setUp() throws Exception {
        File srcDir = new File("src/test/data");
        File testDir = new File("build/testData");
        testDir.mkdirs();

        File srcSSFile = new File(srcDir, "spreadsheet1.xlsx");
        testXLSXFile = new File(testDir, "ss.xlsx");
        if (testXLSXFile.exists()) {
            testXLSXFile.delete();
        }
        Files.copy(srcSSFile.toPath(), testXLSXFile.toPath());
}
    
    

    /**
     * @throws java.lang.Exception
     */
    @AfterEach
    public void tearDown() throws Exception {
        if (testXLSXFile.exists()) {
            testXLSXFile.delete();
        }
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
