package ittimfn.poi.controler.impl;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.apache.poi.openxml4j.util.ZipSecureFile;

import java.io.FileNotFoundException;

import ittimfn.poi.controler.WriteController;

public class SXSSFWriteController extends WriteController {
    private Logger logger = LogManager.getLogger();
    
    public SXSSFWriteController(SXSSFWorkbook workbook, String filepath, String sheetName) {
        this.workbook = workbook;
        logger.info("workbook : {}", this.workbook.getClass());

        this.filepath = filepath;
        logger.info("filepath : {}", this.filepath);

        this.sheetName = sheetName;
        logger.info("sheetName : {}", this.sheetName);
    }
    
    public SXSSFWriteController(SXSSFWorkbook workbook, String filepath, String sheetName, boolean append) {
        this(workbook ,filepath, sheetName);
        
        this.append = append;
        logger.info("append : {}", this.append);
    }
    
    // @Override
    // public void open() throws FileNotFoundException {
    //     // this.baos = new ByteArrayOutputStream();
    //     super.open();
    //     ZipSecureFile.setMinInflateRatio(0.001);
    // }

}
