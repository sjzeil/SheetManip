package edu.odu.cs.sheetManip;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.PrintStream;

//import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.*;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;


public class TestToHTML {

        String outDirName = "build/test";
        File outDir;

        @BeforeEach
        public void setUp() throws Exception {
                outDir = new File(outDirName);
                if (outDir.exists()) {
                        File[] files = outDir.listFiles();
                        assert (files != null);
                        for (File file : files) {
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
        public void testExtractAsHTML() throws IOException {
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                PrintStream strOut = new PrintStream(baos);
                PrintStream oldOut = System.out;
                System.setOut(strOut);

                edu.odu.cs.sheetManip.CLI.toHtml extr = new edu.odu.cs.sheetManip.CLI.toHtml(
                                "src/test/data/spreadsheet1.xlsx",
                                "Main title");
                extr.run();
                System.out.flush();
                System.setOut(oldOut);

                String htmlPage = baos.toString();

                assertThat(htmlPage,
                                containsString("<title>Main title</title>"));
                assertThat(htmlPage,
                                containsString("<h1>Main title</h1>"));
                assertThat(htmlPage,
                                containsString("<h2>in</h2>"));
                assertThat(htmlPage,
                                containsString("<h2>work</h2>"));
                assertThat(htmlPage,
                                containsString("<h2>out</h2>"));

                assertThat(htmlPage, containsString("<table"));
                assertThat(htmlPage, containsString("</table>"));
                assertThat(htmlPage, containsString("<b>Column 2.</b>"));
                assertThat(htmlPage, containsString("<i>Row 4.</i>"));
                assertThat(htmlPage, containsString("<td>9</td>"));

        }

}
