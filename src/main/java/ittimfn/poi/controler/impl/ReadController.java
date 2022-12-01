package ittimfn.poi.controler.impl;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Data;

/**
 * Excelファイルを読み込む。
 * A列のみで、1行目からスペースなく詰まっている前提。
 */
@Data
public class ReadController {
    private Logger logger = LogManager.getLogger();
    
    private XSSFWorkbook workbook;
    private Sheet sheet;

    private String filepath;

    public void open(String filepath) throws IOException {
        this.filepath = filepath;

       logger.info("filepath : {}", this.filepath);

        // ファイルOpen
        this.workbook = new XSSFWorkbook(this.filepath);
        this.sheet = this.workbook.getSheetAt(0);

        logger.info("sheet name : {}", this.sheet.getSheetName());
    }

    public List<String> read() {
        Iterator<Row> rows = this.sheet.rowIterator();
        List<String> returnList = new ArrayList<String>();

        int line = 1;
        while(rows.hasNext()) {
            Cell cell = rows.next().getCell(0);

            // 行数分読み込み。本来は型を確認する必要があるが、今回は文字列前提。
            String cellValue = cell.getStringCellValue();
            logger.info("line : {} , cellValue : {}", line, cellValue);
            
            returnList.add(cellValue);

            line++;
        }

        logger.info("return list size : {}", returnList.size());

        return returnList;

    }

    /**
     * 作成者を取得する。
     */
    public String getCreator() {
        return this.workbook.getProperties().getCoreProperties().getCreator();
    }
    
    /**
     * アプリケーション名を取得する
     */
    public String getProgramName() {
        return this.workbook.getProperties()
                            .getExtendedProperties()
                            .getUnderlyingProperties()
                            .getApplication();
    }

    public void getCustomProperty() {
        // TODO
    }
}
