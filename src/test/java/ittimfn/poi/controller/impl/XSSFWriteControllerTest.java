package ittimfn.poi.controller.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.ArrayList;

import static org.hamcrest.MatcherAssert.*;
import static org.hamcrest.Matchers.*;

import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import ittimfn.poi.controler.impl.ReadController;
import ittimfn.poi.controler.impl.XSSFWriteController;

public class XSSFWriteControllerTest {
    private Logger logger = LogManager.getLogger();

    private XSSFWriteController controller;
    private ReadController reader;

    private static final String EXPORT_HOME
        = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "writetest").toString();

    @BeforeEach
    public void deleteExcelFile() throws IOException {
        FileUtils.cleanDirectory(new File(EXPORT_HOME));
    }

    /**
     * XSSFWorkbookインスタンスを使ってExcelファイル書き込みができることを確認する。
     */
    @Test
    public void writeByXSSFWorkbookTest() throws FileNotFoundException, IOException {
        String filepath = Paths.get(EXPORT_HOME, "byXSSFWorkbook.xlsx").toString();
        XSSFWorkbook workbook = new XSSFWorkbook();
        String sheetName = "test";
        
        List<String> list = Arrays.asList("hoge","piyo");
        List<String> result;
        try {
            this.controller = new XSSFWriteController(workbook, filepath, sheetName);
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

    /**
     * XSSFWorkbookインスタンスを使って、作成者を書き込みできること。
     */
    @Test
    public void writeCreatorBySXSSWorkbook()  throws FileNotFoundException, IOException {
        String filepath = Paths.get(EXPORT_HOME, "writeCreatorBySXSSFWorkbook.xlsx").toString();
        XSSFWorkbook workbook = new XSSFWorkbook();
        String sheetName = "test";

        final String CUSTOM_CREATOR = "customCreator";
        try {
            this.controller = new XSSFWriteController(workbook, filepath, sheetName);
            this.controller.setCreator(CUSTOM_CREATOR);
            this.controller.write(new ArrayList<String>());

            this.reader = new ReadController();
            this.reader.open(filepath);
            assertThat(this.reader.getCreator(), is(CUSTOM_CREATOR));
        } catch (Exception e) {
            logger.error(e.getStackTrace());
            throw e;
        }
    }
    
    /**
     * XSSFWorkbookインスタンスを使って、プログラム名を書き込みできること。
     */
    @Test
    public void writeApplicationNameBySXSSWorkbook()  throws FileNotFoundException, IOException {
        String filepath = Paths.get(EXPORT_HOME, "writeApplicationNameBySXSSFWorkbook.xlsx").toString();
        XSSFWorkbook workbook = new XSSFWorkbook();
        String sheetName = "test";

        final String APPLICATION_NAME = "applicationName";
        try {
            this.controller = new XSSFWriteController(workbook, filepath, sheetName);
            this.controller.setApplicationName(APPLICATION_NAME);
            this.controller.write(new ArrayList<String>());

            this.reader = new ReadController();
            this.reader.open(filepath);
            assertThat(this.reader.getApplicationName(), is(APPLICATION_NAME));
        } catch (Exception e) {
            logger.error(e.getStackTrace());
            throw e;
        }
    }
}