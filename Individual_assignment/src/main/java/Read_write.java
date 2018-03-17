import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.DataFormatter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_write {
     public static void read() throws IOException {
         InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\hp\\Desktop\\RT\\Practicumlist.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

        XSSFWorkbook test = new XSSFWorkbook();

        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row;
        XSSFCell cell;
        
        Iterator rows = sheet.rowIterator();
        Writer writer = null;File file = new File("C:\\Users\\hp\\240450-STIW3054-A172-A1.wiki\\Home.md");
        writer = new BufferedWriter(new FileWriter(file));
        
        boolean num=true;
        DataFormatter data= new DataFormatter();
        while (rows.hasNext()) {
            row = (XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell = (XSSFCell) cells.next();
                String nama = data.formatCellValue(cell);
            
                    writer.write("|");
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
                
                writer.write("--|--|--|--\n");
                System.out.println("--|--|--|--");
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
        }
       public static void gitbash() throws IOException{
    String[] command = {"C:\\Program Files\\Git\\git-bash.exe",
                    };    
        
Runtime.getRuntime().exec(command);
    
    }

      public static void main(String[] args) throws IOException {

        read();
        gitbash();

    }
     }
         
    
    

