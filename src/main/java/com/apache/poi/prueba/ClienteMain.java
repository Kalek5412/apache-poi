package com.apache.poi.prueba;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class ClienteMain {

    public static void main(String[] args) {

        List<Cliente> listaDeClientes = obtenerListado();
        Field[] campos= Cliente.class.getDeclaredFields();

        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja = libro.createSheet("Clientes");

        XSSFCellStyle estiloTitulo = new Styles.Builder().setColorPerzonalizado("C128CE")
                                                         .setTipoPatron(FillPatternType.SOLID_FOREGROUND)
                                                         .setAlineacionHorizontal(HorizontalAlignment.CENTER)
                                                         .build(libro);

        XSSFCellStyle estilosContenido = new Styles.Builder().setColorPerzonalizado("f6ccfa")
                .setTipoPatron(FillPatternType.SOLID_FOREGROUND)
                .setAlineacionHorizontal(HorizontalAlignment.CENTER)
                .setBordeArriba(BorderStyle.THIN)
                .setBordeAbajo(BorderStyle.THIN)
                .setBordeDerecha(BorderStyle.THIN)
                .setBordeIzquierda(BorderStyle.THIN)
                .build(libro);

        XSSFCellStyle estilosFecha = new Styles.Builder().setColorPerzonalizado("f6ccfa")
                .setTipoPatron(FillPatternType.SOLID_FOREGROUND)
                .setAlineacionHorizontal(HorizontalAlignment.CENTER)
                .setBordeArriba(BorderStyle.THIN)
                .setBordeAbajo(BorderStyle.THIN)
                .setBordeDerecha(BorderStyle.THIN)
                .setBordeIzquierda(BorderStyle.THIN)
                .setFormato("dd/MM/yyyy")
                .build(libro);

        XSSFRow fila=null;
        XSSFCell celda=null;

        for(int i=0; i< listaDeClientes.size();i++){
            /* generar la cabecera */
            if(i==0){
                fila=hoja.createRow(0);
                for (int j=0;j< campos.length; j++){
                    celda = fila.createCell(j);
                    celda.setCellValue(campos[j].getName());
                    celda.setCellStyle(estiloTitulo);
                }
            }
            Cliente cliente = listaDeClientes.get(i);
            List<Object> atributos = cliente.obtenerAtributos();
            fila=hoja.createRow(i+1);
            for (int a=0; a<atributos.size(); a++){
                celda= fila.createCell(a);
                if(atributos.get(a) instanceof Long){
                    celda.setCellValue((Long)atributos.get(a));
                    celda.setCellStyle(estilosContenido);
                }
                if(atributos.get(a) instanceof String){
                    celda.setCellValue((String)atributos.get(a));
                    celda.setCellStyle(estilosContenido);
                }
                if(atributos.get(a) instanceof LocalDate){
                    celda.setCellValue((LocalDate)atributos.get(a));
                    celda.setCellStyle(estilosFecha);
                }
                celda.setCellStyle(estilosContenido);
                hoja.autoSizeColumn(a);
            }
        }
        try{
            OutputStream out = new FileOutputStream("C:\\Users\\WANDER\\Documents\\PruebaExcel.xlsx");
            libro.write(out);
            libro.close();
            out.close();
        }catch(Exception e){
            e.printStackTrace();
        }

    }

    public static List<Cliente> obtenerListado(){
        List<Cliente> listCliente = new ArrayList<>();
        listCliente.add(new Cliente(1L, "Santiago", "Pérez", "12345", "santi@admin.com", LocalDate.of(1995, 11, 14)));
        listCliente.add(new Cliente(2L, "María", "Gómez", "67890", "maria.gomez@gmail.com", LocalDate.of(1992, 3, 22)));
        listCliente.add(new Cliente(3L, "Luis", "Fernández", "54321", "luis.fernandez@yahoo.com", LocalDate.of(1988, 7, 9)));
        listCliente.add(new Cliente(4L, "Carla", "Ramírez", "11223", "carla.ramirez@outlook.com", LocalDate.of(1999, 5, 30)));
        listCliente.add(new Cliente(5L, "Jorge", "Torres", "99887", "jorge.torres@hotmail.com", LocalDate.of(1985, 1, 17)));
        listCliente.add(new Cliente(6L, "Lucía", "Vargas", "77665", "lucia.vargas@gmail.com", LocalDate.of(2000, 9, 5)));
        listCliente.add(new Cliente(7L, "Andrés", "Flores", "33445", "andres.flores@empresa.com", LocalDate.of(1993, 12, 1)));
        listCliente.add(new Cliente(8L, "Paola", "Ríos", "22119", "paola.rios@correo.com", LocalDate.of(1997, 2, 8)));
        listCliente.add(new Cliente(9L, "Diego", "Castro", "88990", "diego.castro@gmail.com", LocalDate.of(1989, 10, 25)));
        return  listCliente;
    }
}
