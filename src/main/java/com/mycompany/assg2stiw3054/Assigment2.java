/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.assg2stiw3054;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Writer;
import java.util.Iterator;
import java.util.Scanner;
import java.util.StringTokenizer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author User
 */
public class Assigment2 {

    public static void run() {

        Writer writer = null;
        boolean LineOut = true;

        try {

            DataFormatter dataformat = new DataFormatter();
            FileInputStream excel = new FileInputStream(new File("C:\\Users\\User\\Documents\\NetBeansProjects\\Assg2STIW3054\\Practicum-StudentSupervisorList.xlsx"));
            Workbook workbook = new XSSFWorkbook(excel);
            Sheet data = workbook.getSheetAt(0);
            Iterator<Row> iterator = data.iterator();

            File markdown = new File("C:\\Users\\User\\Documents\\NetBeansProjects\\Assg2STIW3054\\243055.md");
            writer = new BufferedWriter(new FileWriter(markdown));

            while (iterator.hasNext()) {

                Row row = iterator.next();
                Iterator<Cell> cellIterator = row.iterator();

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();
                    String value = dataformat.formatCellValue(cell);

                    System.out.print(value + "|");

                    writer.write(value + "|");

                }
                System.out.println();
                writer.write("\n");
                if (LineOut == true) {
                    writer.write("---|---|---|---|\n");
                    LineOut = false;
                }

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            if (writer != null) {
                writer.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


public static void main(String[] args) throws IOException{
        File file = new File("C:\\Users\\User\\Documents\\NetBeansProjects\\Assg2STIW3054\\243055.md");
            Scanner in = new Scanner(file);
            
            String lines = null;
            int numLine =0;
            int words = 0;
            int chars = 0;

          while(in.hasNextLine())  {
        numLine++;
        lines = in.nextLine();
        chars += lines.length();
        words += new StringTokenizer(lines, " ,").countTokens();
    }

            System.out.println("The total of lines : " + numLine);
            System.out.println("The total of words : " + words);
            System.out.println("The total of characters : " + chars);
}
}