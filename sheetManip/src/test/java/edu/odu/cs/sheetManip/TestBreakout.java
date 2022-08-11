package edu.odu.cs.sheetManip;


import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;

import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.*;

import edu.odu.cs.sheetManip.CLI.BreakOutByRow;
import edu.odu.cs.sheetManip.CLI.ExtractAsCsv;

public class TestBreakout {

	String outDirName = "build/test";
	File outDir;

	@BeforeEach
    public void setUp() throws Exception {
		outDir = new File(outDirName);
    	outDir.mkdirs();
    }

    @AfterEach
    public void tearDown() throws Exception {
        File[] files = outDir.listFiles();
        assert (files != null);
        for (File file: files) {
    		file.delete();
    	}
    	outDir.delete();
    }

    @Test
    public void testBreakOut() throws Exception {
        BreakOutByRow bobr = new BreakOutByRow("src/test/data/sampleCourse.xls", "collectedGrades", 
        		"build/test", "A", "D");
        bobr.run();
        
        assertTrue (new File(outDir, "Row 2..xls").exists());
        assertTrue (new File(outDir, "Row 3..xls").exists());
        assertTrue (new File(outDir, "Row 4..xls").exists());
        assertTrue (new File(outDir, "Sums.xls").exists());
        assertFalse (new File(outDir, ".xls").exists());
        assertFalse (new File(outDir, " .xls").exists());
        
        File ssToCheck = new File(outDir, "Sums.xls");
        File csvToCheck = new File(outDir, "Sums.csv");
        
        ExtractAsCsv contents = new ExtractAsCsv(ssToCheck.getAbsolutePath(), "collectedGrades", csvToCheck.getAbsolutePath());
        contents.run();
        
        BufferedReader in = new BufferedReader(new FileReader(csvToCheck));
        String header = filter(in.readLine());
        assertEquals(",Column2.,Column3.,Column4.", header);
        assertEquals("Sums,3.0,6.0,9.0", filter(in.readLine()));
        assertEquals("Total:,9.0", filter(in.readLine()));
        in.close();
    }

	private String filter(String str) {
		String result = str.replace(" ", "");
		result = result.replace ("\"", "");
		return result;
	}

}
