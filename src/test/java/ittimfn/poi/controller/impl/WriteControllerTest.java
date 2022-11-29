package ittimfn.poi.controller.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import static org.hamcrest.MatcherAssert.*;
import static org.hamcrest.Matchers.*;

import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import ittimfn.poi.controler.impl.ReadController;
import ittimfn.poi.controler.impl.WriteController;

public class WriteControllerTest {
    private Logger logger = LogManager.getLogger();

    private WriteController controller;
    private ReadController reader;

    private static final String EXPORT_HOME
        = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "writetest").toString();

    @BeforeEach
    public void deleteExcelFile() throws IOException {
        FileUtils.cleanDirectory(new File(EXPORT_HOME));
    }

    @Test
    public void writeByXSSFWorkbookTest() throws FileNotFoundException, IOException {
        String filepath = Paths.get(EXPORT_HOME, "byXSSFWorkbook.xlsx").toString();
        XSSFWorkbook workbook = new XSSFWorkbook();
        String sheetName = "test";
        
        List<String> list = Arrays.asList("hoge","piyo");
        List<String> result;
        try {
            this.controller = new WriteController(workbook, filepath, sheetName);
            this.controller.write(list);

            this.reader = new ReadController();
            this.reader.open(filepath);
            result = this.reader.read();
        } catch (Exception e) {
            logger.error(e.getStackTrace());
            throw e;
        }

        assertThat(result.size(), is(list.size()));
        assertThat(result.get(0), is("hoge"));
        assertThat(result.get(1), is("piyo"));
    }

    @Test
    public void writeBySXSSFWorkbookTest() throws FileNotFoundException, IOException {
        String filepath = Paths.get(EXPORT_HOME, "bySXSSFWorkbook.xlsx").toString();
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        String sheetName = "test";
        
        List<String> list = Arrays.asList("hoge","piyo");
        List<String> result;
        try {
            this.controller = new WriteController(workbook, filepath, sheetName);
            this.controller.write(list);

            this.reader = new ReadController();
            this.reader.open(filepath);
            result = this.reader.read();
        } catch (Exception e) {
            logger.error(e.getStackTrace());
            throw e;
        }

        assertThat(result.size(), is(list.size()));
        assertThat(result.get(0), is("hoge"));
        assertThat(result.get(1), is("piyo"));
    }
}
