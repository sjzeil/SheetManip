package edu.odu.cs.sheetManip;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;

import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.*;

import edu.odu.cs.sheetManip.CLI.ExtractAsCsv;
import edu.odu.cs.sheetManip.CLI.SplitByColumn;

public class TestSplitByColumn {

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
	    	outDir.delete();			
		}
    	outDir.mkdirs();
    }

    @AfterEach
    public void tearDown() throws Exception {
    }

    @Test
    public void testSplit() throws Exception {
        SplitByColumn bobr = new SplitByColumn("src/test/data/sampleCourse.xls", "collectedGrades", 
        		"build/test", 1, 6);
        bobr.run();
        
        assertTrue (new File(outDir, "Column 2..xls").exists());
        assertTrue (new File(outDir, "Column 3..xls").exists());
        assertTrue (new File(outDir, "Column 4..xls").exists());
        assertFalse (new File(outDir, ".xls").exists());
        assertFalse (new File(outDir, " .xls").exists());
        
        {
        	File ssToCheck = new File(outDir, "Column 2..xls");
        	File csvToCheck = new File(outDir, "Sums.csv");

        	ExtractAsCsv contents = new ExtractAsCsv(ssToCheck.getAbsolutePath(), 
        			"collectedGrades", csvToCheck.getAbsolutePath());
        	contents.run();

        	BufferedReader in = new BufferedReader(new FileReader(csvToCheck));
        	String header = filter(in.readLine());
        	assertEquals(",Column2.,,", header);
            assertEquals("Row2.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Row3.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Row4.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Sums,3.0,,", filter(in.readLine()));
        	in.close();
        }
        {
        	File ssToCheck = new File(outDir, "Column 3..xls");
        	File csvToCheck = new File(outDir, "Sums.csv");

        	ExtractAsCsv contents = new ExtractAsCsv(ssToCheck.getAbsolutePath(), 
        			"collectedGrades", csvToCheck.getAbsolutePath());
        	contents.run();

        	BufferedReader in = new BufferedReader(new FileReader(csvToCheck));
        	String header = filter(in.readLine());
        	assertEquals(",,Column3.,", header);
            assertEquals("Row2.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Row3.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Row4.,1.0,2.0,3.0", filter(in.readLine()));
        	assertEquals("Sums,,6.0,", filter(in.readLine()));
        	in.close();
        }

    }

    @Test
    public void testSplit_xlsx() throws Exception {
        SplitByColumn bobr = new SplitByColumn("src/test/data/spreadsheet1.xlsx", "out", 
        		"build/test", 1, 6);
        bobr.run();
        
        assertFalse (new File(outDir, "Column 2..xlsx").exists());
        assertTrue (new File(outDir, "Column 3..xlsx").exists());
        assertTrue (new File(outDir, "Column 4..xlsx").exists());
        assertFalse (new File(outDir, ".xls").exists());
        assertFalse (new File(outDir, " .xls").exists());
        
        {
        	File ssToCheck = new File(outDir, "Column 4..xlsx");
        	File csvToCheck = new File(outDir, "Sums.csv");

        	ExtractAsCsv contents = new ExtractAsCsv(ssToCheck.getAbsolutePath(), 
        			"out", csvToCheck.getAbsolutePath());
        	contents.run();

        	BufferedReader in = new BufferedReader(new FileReader(csvToCheck));
        	String header = filter(in.readLine());
        	assertEquals(",,,Column4.", header);
            assertEquals("Row2.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Row3.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Row4.,1.0,2.0,3.0", filter(in.readLine()));
        	assertEquals("Sums,,,9.0", filter(in.readLine()));
        	in.close();
        }
        {
        	File ssToCheck = new File(outDir, "Column 3..xlsx");
        	File csvToCheck = new File(outDir, "Sums.csv");

        	ExtractAsCsv contents = new ExtractAsCsv(ssToCheck.getAbsolutePath(), 
        			"out", csvToCheck.getAbsolutePath());
        	contents.run();

        	BufferedReader in = new BufferedReader(new FileReader(csvToCheck));
        	String header = filter(in.readLine());
        	assertEquals(",,Column3.,", header);
            assertEquals("Row2.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Row3.,1.0,2.0,3.0", filter(in.readLine()));
            assertEquals("Row4.,1.0,2.0,3.0", filter(in.readLine()));
        	assertEquals("Sums,,6.0,", filter(in.readLine()));
        	in.close();
        }

    }

    private String filter(String str) {
		String result = str.replace(" ", "");
		result = result.replace ("\"", "");
		return result;
	}

}
