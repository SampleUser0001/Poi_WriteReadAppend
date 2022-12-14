package ittimfn.poi.controler.impl;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFWriteController {
    private Logger logger = LogManager.getLogger();
    
    private XSSFWorkbook workbook;
    private String sheetName;

    private String filepath;

    public XSSFWriteController(XSSFWorkbook workbook, String filepath, String sheetName) {
        this.workbook = workbook;
        logger.info("workbook : {}", this.workbook.getClass());

        this.filepath = filepath;
        logger.info("filepath : {}", this.filepath);

        this.sheetName = sheetName;
        logger.info("sheetName : {}", this.sheetName);
    }

    public void write(List<String> list) throws FileNotFoundException, IOException {
        try {
            Sheet sheet = this.workbook.createSheet(this.sheetName);

            for(int rowIndex = 0 ; rowIndex < list.size() ; rowIndex++) {
                Row row = sheet.createRow(rowIndex);
                Cell cell = row.createCell(0);

                String value = list.get(rowIndex);
                logger.info("line : {} , value : {}", rowIndex+1, value);
                cell.setCellValue(value);
            }

            try(FileOutputStream out = new FileOutputStream(this.filepath)) {
                this.workbook.write(out);
                this.workbook.close();
                out.close();
            }
        } catch(Exception e) {
            logger.error(e);
            e.printStackTrace();
            throw e;
        }
    }
    
    /**
     * workbookに作成者を書き込む
     */
    public void setCreator(String creator) {
        this.workbook.getProperties().getCoreProperties().setCreator(creator);
    }

}
