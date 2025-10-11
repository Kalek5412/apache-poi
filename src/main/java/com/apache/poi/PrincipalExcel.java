package com.apache.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class PrincipalExcel {
    public static void main(String[] args) {
        /*
        * Crear libro
        * Workbook libro = new XSSFWorkbook(); //.xlsx 2007-2025
        * Workbook libro = new HSSFWorkbook(); //.xls  1997-2003
        *
        */

        Workbook libro = new XSSFWorkbook();
        Sheet hoja = libro.createSheet("Personas");
        Row fila = hoja.createRow(2);
        Cell nombre = fila.createCell(1);
        Cell edad = fila.createCell(2);
        Cell ciudad = fila.createCell(3);

        nombre.setCellValue("Nombre");
        edad.setCellValue("Edad");
        ciudad.setCellValue("Ciudad");


        try{
            OutputStream out = new FileOutputStream("C:\\Users\\WANDER\\Documents\\ArchivoExcel.xlsx");
            libro.write(out);
            libro.close();
            out.close();
        }catch(Exception e){
            e.printStackTrace();
        }

    }
}