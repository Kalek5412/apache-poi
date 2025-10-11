package com.apache.poi;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class ColoresExcel {
    public static void main(String[] args) {

        XSSFColor verClaro=crearColor("61F744");

        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja = libro.createSheet("Personas");
        XSSFRow fila = hoja.createRow(1);
        /*crear celdas*/
        XSSFCell celda = fila.createCell(1);
        XSSFCellStyle celdaStyle = libro.createCellStyle();
        XSSFCell celda2 = fila.createCell(2);
        XSSFCellStyle celdaStyle2 = libro.createCellStyle();

        /*configuracion deestilos*/
        celdaStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        celdaStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        celdaStyle2.setFillForegroundColor(verClaro);
        celdaStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        /*border*/
        celdaStyle.setBorderBottom(BorderStyle.THIN);
        celdaStyle.setBorderTop(BorderStyle.THIN);
        celdaStyle.setBorderLeft(BorderStyle.THIN);
        celdaStyle.setBorderRight(BorderStyle.THIN);

        /*configuracion de celda*/
        celda.setCellValue("Estilo predeterminado");
        celda.setCellStyle(celdaStyle);

        celda2.setCellValue("Estilo perzonalixado");
        celda2.setCellStyle(celdaStyle2);

        hoja.autoSizeColumn(1);
        hoja.autoSizeColumn(2);

        try{
            OutputStream out = new FileOutputStream("C:\\Users\\WANDER\\Documents\\EstilosExcel.xlsx");
            libro.write(out);
            libro.close();
            out.close();
        }catch(Exception e){
            e.printStackTrace();
        }
    }

    public static  XSSFColor crearColor(String colorHex){
        try{
            byte[] rgb = Hex.decodeHex(colorHex);
            return new XSSFColor(rgb);
        }catch(DecoderException e){
            e.printStackTrace();
            throw new RuntimeException("Error al crear color");
        }
    }
}

