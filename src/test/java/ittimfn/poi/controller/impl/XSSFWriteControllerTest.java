package ittimfn.poi.controller.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.io.FileOutputStream;

import static org.hamcrest.MatcherAssert.*;
import static org.hamcrest.Matchers.*;

import java.io.FileInputStream;

import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import org.apache.poi.openxml4j.util.ZipSecureFile;

import ittimfn.poi.controler.impl.ReadController;
import ittimfn.poi.controler.impl.XSSFWriteController;
import ittimfn.poi.controler.impl.SXSSFWriteController;

public class XSSFWriteControllerTest {
    private Logger logger = LogManager.getLogger();

    private XSSFWriteController controller;
    private SXSSFWriteController sxssfController;

    private ReadController reader;

    private static final String EXPORT_HOME
        = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "writetest").toString();
    private static final String FOR_DOWNLOAD
        = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "download").toString();

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
            this.controller.open();
            this.controller.createSheet();
            this.controller.write(list);
            this.controller.writeToWorkbook();
            this.controller.close();

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

    /**
     * XSSFWorkbookインスタンスを使って、作成者を書き込みできること。
     */
    @Test
    public void writeCreatorBySXSSWorkbook() throws FileNotFoundException, IOException {
        String filepath = Paths.get(EXPORT_HOME, "writeCreatorBySXSSFWorkbook.xlsx").toString();
        XSSFWorkbook workbook = new XSSFWorkbook();
        String sheetName = "test";

        final String CUSTOM_CREATOR = "customCreator";
        try {
            this.controller = new XSSFWriteController(workbook, filepath, sheetName);
            this.controller.open();
            this.controller.setCreator(CUSTOM_CREATOR);
            this.controller.writeToWorkbook();
            this.controller.close();

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
    public void writeApplicationNameBySXSSWorkbook() throws FileNotFoundException, IOException {
        String filepath = Paths.get(EXPORT_HOME, "writeApplicationNameBySXSSFWorkbook.xlsx").toString();
        XSSFWorkbook workbook = new XSSFWorkbook();
        String sheetName = "test";

        final String APPLICATION_NAME = "applicationName";
        try {
            this.controller = new XSSFWriteController(workbook, filepath, sheetName);
            this.controller.open();
            this.controller.setApplicationName(APPLICATION_NAME);
            this.controller.writeToWorkbook();
            this.controller.close();

            this.reader = new ReadController();
            this.reader.open(filepath);
            assertThat(this.reader.getApplicationName(), is(APPLICATION_NAME));
        } catch (Exception e) {
            logger.error(e.getStackTrace());
            throw e;
        }
    }

    /**
     * SXSSFでセルに書き込み、XSSFでプロパティに書き込む。
     */
    @Test
    public void bothUseTest() throws FileNotFoundException, IOException {
        logger.info("bothUseTest start.");
        String filepath = Paths.get(FOR_DOWNLOAD, "bothWorkbook.xlsx").toString();
        String sheetName = "test";

        List<String> list = Arrays.asList("hoge","piyo");
        List<String> result;

        final String CUSTOM_CREATOR = "customCreator";
        final String APPLICATION_NAME = "applicationName";

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();

        try {

            this.controller = new XSSFWriteController(xssfWorkbook, filepath, sheetName);
            this.controller.open();
            this.controller.setCreator(CUSTOM_CREATOR);
            this.controller.setApplicationName(APPLICATION_NAME);
            this.controller.writeToWorkbook();
            this.controller.workbookClose();
            this.controller.close();

            xssfWorkbook = new XSSFWorkbook(new FileInputStream(filepath));
            sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook);
            
            this.sxssfController = new SXSSFWriteController(sxssfWorkbook, filepath, sheetName);
            this.sxssfController.open();
            this.sxssfController.createSheet();
            ZipSecureFile.setMinInflateRatio(0.001);
            this.sxssfController.write(list);
            this.sxssfController.writeToWorkbook();
            this.sxssfController.workbookClose();
            this.sxssfController.close();

            this.reader = new ReadController();
            this.reader.open(filepath);

            this.reader.openSheet();
            result = this.reader.read();

        } catch (Exception e) {
            e.printStackTrace();
            logger.error(e.getStackTrace());
            throw e;
        }

        assertThat(result.size(), is(list.size()));
        assertThat(result.get(0), is("hoge"));
        assertThat(result.get(1), is("piyo"));

        assertThat(this.reader.getCreator(), is(CUSTOM_CREATOR));
        assertThat(this.reader.getApplicationName(), is(APPLICATION_NAME));

   }

}