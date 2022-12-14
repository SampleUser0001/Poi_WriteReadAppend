package ittimfn.poi.controler.impl;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ittimfn.poi.controler.WriteController;

public class XSSFWriteController extends WriteController {
    private Logger logger = LogManager.getLogger();
    
    public XSSFWriteController(XSSFWorkbook workbook, String filepath, String sheetName) {
        this.workbook = workbook;
        logger.info("workbook : {}", this.workbook.getClass());

        this.filepath = filepath;
        logger.info("filepath : {}", this.filepath);

        this.sheetName = sheetName;
        logger.info("sheetName : {}", this.sheetName);
    }
    
    public XSSFWriteController(XSSFWorkbook workbook, String filepath, String sheetName, boolean append) {
        this(workbook, filepath, sheetName);
        
        this.append = append;
        logger.info("append : {}" , this.append);
    }
    
    
    /**
     * workbookに作成者を書き込む
     */
    public void setCreator(String creator) {
        XSSFWorkbook xssf = (XSSFWorkbook)this.workbook;
        xssf.getProperties().getCoreProperties().setCreator(creator);
    }
    
    /**
     * workbookにプログラム名を書き込む
     */
     
    public void setApplicationName(String applicationName) {
        XSSFWorkbook xssf = (XSSFWorkbook)this.workbook;
        xssf.getProperties()
            .getExtendedProperties()
            .getUnderlyingProperties()
            .setApplication(applicationName);
    }

}
