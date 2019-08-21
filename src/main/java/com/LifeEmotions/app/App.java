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
            private static  String XLS_FILE_NAME;
            private static  String CSV_FILE_NAME;


            public static void main( String[] args ){

                if (args.length!=2){
                    System.out.println();
                    System.out.println("###################################################################################################");
                    System.out.println("Error running the CSV generator!");
                    System.out.println("Please copy the XML file to JAR directory. Then try to run:");
                    System.out.println("");
                    System.out.println("java -jar ETSgroupAddressCSVgenetartor-1.0-jar-with-dependencies.jar [XML_file_name].xlsm [CSV_file_name_to_generate].csv");
                    System.out.println();
                    System.out.println("Note that you can use te command passing paths for the files as argument.\nIt will allows you to access and generate files in different directories.");
                    System.out.println("###################################################################################################");
                    System.out.println();

                    return;
                }

                XLS_FILE_NAME = args[0];
                CSV_FILE_NAME = args[1];

                try {

                    System.out.println("XML File is Readable - "+ new File(XLS_FILE_NAME).canRead() );

                    PrintWriter writer = new PrintWriter(new File(CSV_FILE_NAME),"UTF-8");
                    StringBuilder sb;

                    FileInputStream excelFile = new FileInputStream(new File(XLS_FILE_NAME));
                    Workbook workbook = new XSSFWorkbook(excelFile);

                    Iterator<Sheet> sheetIterator = workbook.sheetIterator();
                    System.out.println("\n######################## Parsing file " + XLS_FILE_NAME + "########################\n");

                    while (sheetIterator.hasNext()) {

                        Sheet currentSheet = sheetIterator.next();
                        System.out.println("\n############ Parsing sheet " +": "+ currentSheet.getSheetName()+" ############\n");

                        Iterator<Row> rowIterator = currentSheet.iterator();

                        while (rowIterator.hasNext()) {

                            Row currentRow = rowIterator.next();

                            System.out.println(">> Row Number: "+ currentRow.getRowNum()+" processed..");
                            //System.out.println(">> Last cell: " +currentRow.getLastCellNum());



                            if (currentRow.getLastCellNum()>5 && !currentRow.getCell(5).getStringCellValue().isEmpty()){

                                sb=new StringBuilder();

                                sb.append("\" \",");
                                sb.append("\" \",");

                                String cell_C = currentRow.getCell(2).getStringCellValue();
                                String cell_D = currentRow.getCell(3).getStringCellValue();

                                if (!cell_C.isEmpty() && !cell_D.isEmpty()){
                                    sb.append('\"' + cell_C +" - " +cell_D +'\"' + ",");
                                }else if (cell_C.isEmpty() && !cell_D.isEmpty()){
                                    sb.append('\"' + cell_D +'\"' + ",");
                                }else if (!cell_C.isEmpty() && cell_D.isEmpty()){
                                    sb.append('\"' + cell_C +'\"' + ",");
                                }else if (cell_C.isEmpty() && cell_D.isEmpty()){
                                    sb.append('\"' + "EMPTY" +'\"' + ",");
                                }

                                String cell_F = currentRow.getCell(5).getStringCellValue();
                                sb.append('\"' + cell_F +'\"'+ ",");
                                sb.append("\" \",");
                                sb.append("\" \",");
                                sb.append("\" \",");
                                sb.append("\" \",");
                                sb.append("\"Auto\"");

                                System.out.println(">> Generate CSV line: " + sb.toString());
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
