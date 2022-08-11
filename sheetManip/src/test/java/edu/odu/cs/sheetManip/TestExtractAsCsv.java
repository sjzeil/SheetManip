package edu.odu.cs.sheetManip;


import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.*;

import edu.odu.cs.sheetManip.CLI.ExtractAsCsv;

public class TestExtractAsCsv {

	String outDirName = "build/test";
	File outDir;

	@BeforeEach
    public void setUp() throws Exception {
		outDir = new File(outDirName);
		if (outDir.exists()) {
		    File[] files = outDir.listFiles();
		    assert (files != null);
	    	for (File file: files) {
	    		file.delete();
	    	}
		} else {
			outDir.mkdirs();
		}
    }

    @AfterEach
    public void tearDown() throws Exception {
    }


    @Test
    public void testExtractAsCsv_xls() throws IOException {
        File csvToCheck = new File(outDir, "extr1.csv");
        ExtractAsCsv extr = new ExtractAsCsv("src/test/data/spreadsheet1.xls", 
                "out", csvToCheck.getCanonicalPath());
        extr.run();
        
        
        BufferedReader in = new BufferedReader(new FileReader(csvToCheck));
        String header = filter(in.readLine());
        assertEquals(",Column2.,Column3.,Column4.", header);
        assertEquals("Row2.,1.0,2.0,3.0", filter(in.readLine()));
        assertEquals("Row3.,1.0,2.0,3.0", filter(in.readLine()));
        assertEquals("Row4.,1.0,2.0,3.0", filter(in.readLine()));
        assertEquals("Sums,3.0,6.0,9.0", filter(in.readLine()));
        in.close();
    }
    
    @Test
    public void testExtractAsCsv_xlsx() throws IOException {
        File csvToCheck = new File(outDir, "extr2.csv");
        ExtractAsCsv extr = new ExtractAsCsv("src/test/data/spreadsheet1.xlsx", 
                "out", csvToCheck.getCanonicalPath());
        extr.run();
        
        
        BufferedReader in = new BufferedReader(new FileReader(csvToCheck));
        String header = filter(in.readLine());
        assertEquals(",Column2.,Column3.,Column4.", header);
        assertEquals("Row2.,1.0,2.0,3.0", filter(in.readLine()));
        assertEquals("Row3.,1.0,2.0,3.0", filter(in.readLine()));
        assertEquals("Row4.,1.0,2.0,3.0", filter(in.readLine()));
        assertEquals("Sums,NAN,6.0,9.0", filter(in.readLine()));
        in.close();
    }

    private String filter(String str) {
		String result = str.replace(" ", "");
		result = result.replace ("\"", "");
		return result;
	}


}
