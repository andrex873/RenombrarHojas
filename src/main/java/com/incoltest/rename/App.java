package com.incoltest.rename;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

/**
 * Hello world!
 *
 */
public class App 
{
    
    public static void main( String[] args ) throws IOException, Exception
    {
        
        if(args.length != 1){
            throw new Exception("Debe enviar la ruta del archivo como argumento");
        }
        int hojaIdx = 3;
        File file = new File(args[0]);
        HSSFWorkbook libro = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet nombres = libro.getSheet("NOMBRES");
        Iterator<Row> iterator = nombres.iterator();
        while(iterator.hasNext()){
            Row row = iterator.next();
            String hojaNombre = row.getCell(0).getStringCellValue();
            String hojaAdmin = "-";
            HSSFSheet nuevaHoja = libro.cloneSheet(2);
            nuevaHoja.getRow(2).getCell(9).setCellValue(hojaNombre);
            nuevaHoja.getRow(44).getCell(9).setCellValue(hojaNombre);
            nuevaHoja.getRow(4).getCell(2).setCellValue(hojaAdmin);
            nuevaHoja.getRow(46).getCell(2).setCellValue(hojaAdmin);
            libro.setSheetName(hojaIdx++, hojaNombre);
        }
        libro.write(new FileOutputStream(file));
    }
}
