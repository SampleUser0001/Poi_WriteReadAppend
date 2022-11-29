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

    @Test
    public void readSingleLineExcel() throws IOException {
        String filepath = Paths.get(System.getProperty("user.dir"), "src", "test", "resources", "singleLine.xlsx").toString();
        this.controller.open(filepath);

        List<String> resultList = this.controller.read();

        assertThat(resultList.size(), is(1));
        assertThat(resultList.get(0), is(equalTo("hogehoge")));
    }
    
}
