/**
 * 
 */
package edu.odu.cs.sheetManip;


import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.*;

/**
 * @author zeil
 *
 */
public class TestInjectProperties {
    
    File testXLSXFile;
    
    Properties properties;
    
    
    /**
     * @throws java.lang.Exception
     */
    @BeforeEach
    public void setUp() throws Exception {
        File srcDir = new File("src/test/data");
        File testDir = new File("build/testData");
        testDir.mkdirs();

        File srcSSFile = new File(srcDir, "sheetWithProperties.xlsx");
        testXLSXFile = new File(testDir, "ss.xlsx");
        if (testXLSXFile.exists()) {
            testXLSXFile.delete();
        }
        Files.copy(srcSSFile.toPath(), testXLSXFile.toPath());

        properties = new Properties();
        properties.setProperty("inb2", "10");
        properties.setProperty("inb3", "20");
        properties.setProperty("inb4", "30");
        properties.setProperty("inc2", "1");
        properties.setProperty("inc3", "2");
        properties.setProperty("inc4", "4");
        properties.setProperty("headerB", "Column 2");
        
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
     * Test method for {@link edu.odu.cs.sheetManip.Spreadsheet#injectProperties(java.util.Properties, java.lang.String, java.lang.String)}.
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */
    @Test
    public void testInjection1() throws EncryptedDocumentException, InvalidFormatException, IOException {
        Spreadsheet ss = new Spreadsheet(testXLSXFile);
        try {
            ss.injectProperties (properties, "{{", "}}");
            assertEquals ("Column 2", ss.getCellValue("in", 0, 1));
            assertEquals ("10.0", ss.getCellValue("in", 1, 1));
            assertEquals ("20.0", ss.getCellValue("in", 2, 1));
            assertEquals ("30.0", ss.getCellValue("in", 3, 1));

            assertEquals ("Column 2.", ss.getCellValue("out", 0, 1));
            assertEquals ("11.0", ss.getCellValue("out", 1, 1));
            assertEquals ("21.0", ss.getCellValue("out", 2, 1));
            assertEquals ("31.0", ss.getCellValue("out", 3, 1));
            assertEquals ("63.0", ss.getCellValue("out", 5, 1));
        } finally {
            ss.close();
        }
    }

    /**
    * Test method for {@link edu.odu.cs.sheetManip.Spreadsheet#injectProperties(java.util.Properties, java.lang.String, java.lang.String)}.
    * @throws IOException 
    * @throws InvalidFormatException 
    * @throws EncryptedDocumentException 
    */
   @Test
   public void testInjection2() throws EncryptedDocumentException, InvalidFormatException, IOException {
       Spreadsheet ss = new Spreadsheet(testXLSXFile);
       try {
           ss.injectProperties (properties, "[[", "]]");
           assertEquals ("{{headerB}}", ss.getCellValue("in", 0, 1));
           assertEquals ("1.0", ss.getCellValue("in", 1, 2));
           assertEquals ("2.0", ss.getCellValue("in", 2, 2));
           assertEquals ("4.0", ss.getCellValue("in", 3, 2));

           assertEquals ("3.0", ss.getCellValue("out", 1, 2));
           assertEquals ("4.0", ss.getCellValue("out", 2, 2));
           assertEquals ("6.0", ss.getCellValue("out", 3, 2));
           assertEquals ("13.0", ss.getCellValue("out", 5, 2));
       } finally {       
           ss.close();
       }
   }

}
