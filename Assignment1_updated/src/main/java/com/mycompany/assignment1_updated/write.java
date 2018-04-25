package com.mycompany.assignment1_updated;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Writer;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class write implements readWrite {

    @Override
    public void read() {
    }

    @Override
    public void writeFile() {
        InputStream ExcelFileToRead = null;
        try{
            ExcelFileToRead = new FileInputStream("C:\\Users\\hp\\Desktop\\RT\\Practicumlist.xlsx");
            XSSFWorkbook wb = null;
            try {
                wb = new XSSFWorkbook(ExcelFileToRead);
            } catch (IOException ex) {
                Logger.getLogger(write.class.getName()).log(Level.SEVERE, null, ex);
            }
            XSSFWorkbook test = new XSSFWorkbook();
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;
            Iterator rows = sheet.rowIterator();
            Writer writer = null;
            File file = new File("C:\\Users\\hp\\240450-STIW3054-A172-A1.wiki\\Home2.md");
            writer = new BufferedWriter(new FileWriter(file));
            boolean num=true;
            DataFormatter data= new DataFormatter();
            while (rows.hasNext()) {
                row = (XSSFRow) rows.next();
                Iterator cells = row.cellIterator();
                while (cells.hasNext()) {
                    cell = (XSSFCell) cells.next();
                    String nama = data.formatCellValue(cell);
                    writer.write("");
                    // writer.write("|");
                    if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        
                        System.out.print(nama + " ");
                        
                        writer.write(nama + " ");
                    } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                        System.out.print(nama + " ");
                        writer.write(nama + " ");
                    } else {
                        System.out.print(" ");
                    }
                }
                
                System.out.println();
                writer.write("\n");
                if (num==true){
                    writer.write("");
                    //writer.write("--|--|--|--\n");
                    System.out.println("");
                    // System.out.println("--|--|--|--");
                    num=false;
                }
                
                
            }
            try{
                if (writer!=null){
                    writer.close();
                    
                }
            }catch(IOException e){
                e.printStackTrace();
            }
        }catch(FileNotFoundException ex){
                Logger.getLogger(write.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
            Logger.getLogger(write.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                ExcelFileToRead.close();
            } catch (IOException ex) {
                Logger.getLogger(write.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        }

}
