package com.LifeEmotions.app;
//package com.mkyong;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

        public class App
        {
            private static  String XLS_FILE_NAME= "C:\\Users\\nuno.martins\\IdeaProjects\\ETSgroupAddressCSVgenetartor\\ETSgroupAddressCSVgenetartor\\files\\Group Address Generator Inputs.xlsx";
            private static  String CSV_FILE_NAME = "C:\\Users\\nuno.martins\\IdeaProjects\\ETSgroupAddressCSVgenetartor\\ETSgroupAddressCSVgenetartor\\files\\test.csv";


            public static void main( String[] args ){

                //XLS_FILE_NAME = args[0];
                //CSV_FILE_NAME = args[1];

                try {

                    PrintWriter writer = new PrintWriter(new File(CSV_FILE_NAME));
                    StringBuilder sb;

                    FileInputStream excelFile = new FileInputStream(new File(XLS_FILE_NAME));
                    Workbook workbook = new XSSFWorkbook(excelFile);

                    Iterator<Sheet> sheetIterator = workbook.sheetIterator();
                    System.out.println("\n######################## Parsing file " + XLS_FILE_NAME+ "########################\n");

                    while (sheetIterator.hasNext()) {

                        Sheet currentSheet = sheetIterator.next();
                        System.out.println("\n############ Parsing sheet " +": "+ currentSheet.getSheetName()+" ############\n");

                        Iterator<Row> rowIterator = currentSheet.iterator();

                        while (rowIterator.hasNext()) {

                            Row currentRow = rowIterator.next();

                            if (currentRow.getLastCellNum()>5 && !currentRow.getCell(5).getStringCellValue().isEmpty()){

                                sb=new StringBuilder();

                                sb.append("\" \",");
                                sb.append("\" \",");

                                String cell_C = currentRow.getCell(2).getStringCellValue();
                                if (cell_C.isEmpty()){
                                    cell_C="EMPTY";
                                }
                                sb.append('\"' + cell_C +'\"' + ",");

                                String cell_F = currentRow.getCell(5).getStringCellValue();
                                sb.append('\"' + cell_F +'\"'+ ",");
                                sb.append("\" \",");
                                sb.append("\" \",");
                                sb.append("\" \",");
                                sb.append("\" \",");
                                sb.append("\"Auto\"");

                                System.out.println(sb.toString());
                                writer.write(sb.append('\n').toString());
                         }
                    }
            }
            writer.close();
            System.out.println("\n######################## Parsing Successfully ########################\n");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        catch (IOException e) {
            e.printStackTrace();
        }

    }

}
