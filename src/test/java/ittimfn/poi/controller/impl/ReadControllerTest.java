package ittimfn.poi.controller.impl;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.List;

import static org.hamcrest.MatcherAssert.*;
import static org.hamcrest.Matchers.*;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import ittimfn.poi.controler.impl.ReadController;

public class ReadControllerTest {
    private ReadController controller;

    @BeforeEach
    public void setUp() {
        this.controller = new ReadController();
    }

    /**
     * ファイルの読み込みができることを確認する。
     */
    @Test
    public void readSingleLineExcel() throws IOException {
        String filepath = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "singleLine.xlsx").toString();
        this.controller.open(filepath);

        List<String> resultList = this.controller.read();

        assertThat(resultList.size(), is(1));
        assertThat(resultList.get(0), is(equalTo("hogehoge")));
    }
    
    /**
     * 作成者が取得できることを確認する。
     */
    @Test
    public void getCreatorTest() throws IOException {
        String filepath = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "Poi_properties_read_test.xlsx").toString();
        this.controller.open(filepath);

        String creator = this.controller.getCreator();
        
        assertThat(creator, is(equalTo("hogehoge")));
        
    }

    /**
     * プログラム名が取得できることを確認する。
     */
    @Test
    public void getProgramNameTest() throws IOException {
        String filepath = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "Poi_properties_read_test.xlsx").toString();
        this.controller.open(filepath);

        String programName = this.controller.getProgramName();
        
        assertThat(programName, is(equalTo("Microsoft Excel")));
        
    }
    
}
