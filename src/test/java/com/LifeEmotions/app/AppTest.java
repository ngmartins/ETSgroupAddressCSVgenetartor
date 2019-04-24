package com.LifeEmotions.app;

import org.junit.Test;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

import static org.junit.Assert.assertEquals;

/**
 * Unit test for simple App.
 */
public class AppTest
{
    private   String XLS_FILE_NAME = "C:\\Users\\nuno.martins\\IdeaProjects\\ETSgroupAddressCSVgenetartor\\ETSgroupAddressCSVgenetartor\\testFiles\\XMLSampleFile.xlsm";
    private   String CSV_FILE_NAME_GENERATED = "C:\\Users\\nuno.martins\\IdeaProjects\\ETSgroupAddressCSVgenetartor\\ETSgroupAddressCSVgenetartor\\testFiles\\CSVGeneratedFile.csv";

    private   String CSV_FILE_TO_COMPARE = "C:\\Users\\nuno.martins\\IdeaProjects\\ETSgroupAddressCSVgenetartor\\ETSgroupAddressCSVgenetartor\\testFiles\\CSVComparsionFile.csv";


    @Test
    public void shouldAnswerWithTrue() throws IOException {

        String[] args = {XLS_FILE_NAME,CSV_FILE_NAME_GENERATED};
        App.main(args);


        byte[] ResultBytes = Files.readAllBytes(Paths.get(CSV_FILE_NAME_GENERATED));
        byte[] ExpectedBytes = Files.readAllBytes(Paths.get(CSV_FILE_TO_COMPARE));

        String Result = new String(ResultBytes, StandardCharsets.UTF_8);
        String Expected = new String(ExpectedBytes, StandardCharsets.UTF_8);

        assertEquals("The content in the strings should match", Result, Expected);
    }


}
