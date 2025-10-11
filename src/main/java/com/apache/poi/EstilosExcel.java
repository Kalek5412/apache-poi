package com.apache.poi;


import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class EstilosExcel {

    public static void main(String[] args) {

        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja = libro.createSheet("Personas");
        XSSFRow fila = hoja.createRow(1);
        XSSFCell celda = fila.createCell(1);
        /*config celda*/
        XSSFCellStyle celdaStyle = libro.createCellStyle();

        celdaStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        celdaStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        celdaStyle.setBorderBottom(BorderStyle.THIN);
        celdaStyle.setBorderTop(BorderStyle.THIN);
        celdaStyle.setBorderLeft(BorderStyle.THIN);
        celdaStyle.setBorderRight(BorderStyle.THIN);

        celda.setCellValue("Estilo socn apache poi");
        celda.setCellStyle(celdaStyle);

        hoja.autoSizeColumn(1);

        try{
            OutputStream out = new FileOutputStream("C:\\Users\\WANDER\\Documents\\EstilosExcel.xlsx");
            libro.write(out);
            libro.close();
            out.close();
        }catch(Exception e){
            e.printStackTrace();
        }
    }


}
