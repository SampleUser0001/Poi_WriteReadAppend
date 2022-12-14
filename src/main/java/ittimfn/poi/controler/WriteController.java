package ittimfn.poi.controler;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import lombok.Data;

@Data
public abstract class WriteController {
    private Logger logger = LogManager.getLogger();
    
    protected Workbook workbook;
    protected Sheet sheet;
    protected String sheetName;
    protected String filepath;
    
    protected int rowIndex = 0;
    
    protected boolean append = false;
    
    protected FileOutputStream outputStream;

    public void createSheet() {
        int sheetIndex = this.workbook.getSheetIndex(this.sheetName);
        if(sheetIndex != -1) {
            this.workbook.removeSheetAt(sheetIndex);
        }
        this.sheet = this.workbook.createSheet(this.sheetName);
    }
    public void openSheet() {
        this.sheet = this.workbook.getSheet(this.sheetName);
    }

    public void open() throws FileNotFoundException {
        this.outputStream = new FileOutputStream(this.filepath, this.append);
    }
    
    public void workbookClose() throws IOException {
        this.workbook.close();
    }
    
    public void close() throws IOException {
        this.outputStream.close();
    }
    
    public void writeToWorkbook() throws IOException {
        this.workbook.write(this.outputStream);
//        this.outputStream.flush();
    }

    public void write(List<String> list) throws FileNotFoundException, IOException {
        for(; this.rowIndex < list.size() ; this.rowIndex++) {
            Row row = this.sheet.createRow(this.rowIndex);
            Cell cell = row.createCell(0);

            String value = list.get(this.rowIndex);
            logger.info("line : {} , value : {}", this.rowIndex+1, value);
            cell.setCellValue(value);
        }
    }
}
