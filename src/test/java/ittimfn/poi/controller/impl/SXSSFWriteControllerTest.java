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
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import ittimfn.poi.controler.impl.ReadController;
import ittimfn.poi.controler.impl.SXSSFWriteController;

public class SXSSFWriteControllerTest {
    private Logger logger = LogManager.getLogger();

    private SXSSFWriteController controller;
    private ReadController reader;

    private static final String EXPORT_HOME
        = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "writetest").toString();
    private static final String RESOURCES
        = Paths.get(System.getProperty("user.dir"), "src", "test", "resources").toString();
    
    @BeforeEach
    public void deleteExcelFile() throws IOException {
        FileUtils.cleanDirectory(new File(EXPORT_HOME));
    }

    /**
     * SXSSFWorkbookインスタンスを使ってExcelファイル書き込みができることを確認する。
     */
    @Test
    public void writeBySXSSFWorkbookTest() throws FileNotFoundException, IOException {
        String filepath = Paths.get(EXPORT_HOME, "SXSSFWriteOrigin.xlsx").toString();

        String sheetName = "test";
        
        List<String> list = Arrays.asList("hoge","piyo");
        List<String> result;
        try {
            SXSSFWorkbook workbook = new SXSSFWorkbook();

            this.controller = new SXSSFWriteController(workbook, filepath, sheetName);
            this.controller.createSheet();
            this.controller.open();
            this.controller.write(list);
            this.controller.writeToWorkbook();
            this.controller.close();
            this.controller.workbookClose();

            this.reader = new ReadController();
            this.reader.open(filepath);
            this.reader.openSheet();
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