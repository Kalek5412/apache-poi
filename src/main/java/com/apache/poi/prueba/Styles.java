package com.apache.poi.prueba;

import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Styles {
    public static class Builder {
        private short colorDefecto;
        private XSSFColor colorPerzonalizado;
        private FillPatternType tipoPatron;
        private XSSFFont fuente;
        private String formato;
        private HorizontalAlignment alineHorizontal;
        private VerticalAlignment alineVertical;
        private BorderStyle borArriba;
        private BorderStyle borAbajo;
        private BorderStyle borDerecho;
        private BorderStyle borIzquierdo;

        public Builder setColorDefecto(short colorDefecto){
            this.colorDefecto=colorDefecto;
            return this;
        }

        public Builder setColorPerzonalizado(String colorPerzonalizado){
          try{
              byte[] rgb= Hex.decodeHex(colorPerzonalizado);
              this.colorPerzonalizado = new XSSFColor(rgb);
              return this;
          }catch(Exception e){
              throw  new RuntimeException("Error l decodificar le color.");
          }
        }

        public Builder setTipoPatron(FillPatternType tipoPatron){
            this.tipoPatron=tipoPatron;
            return this;
        }

        public Builder setFuente(XSSFFont fuente){
            this.fuente = fuente;
            return this;
        }
        public Builder setFormato(String formato){
            this.formato = formato;
            return this;
        }
        public Builder setAlineacionHorizontal(HorizontalAlignment alineHorizontal){
            this.alineHorizontal = alineHorizontal;
            return this;
        }
        public Builder setAlineacionVertical(VerticalAlignment alineVertical){
            this.alineVertical = alineVertical;
            return this;
        }

        public Builder setBordeArriba(BorderStyle borArriba){
            this.borArriba = borArriba;
            return this;
        }
        public Builder setBordeAbajo(BorderStyle borAbajo){
            this.borAbajo = borAbajo;
            return this;
        }      public Builder setBordeDerecha(BorderStyle borDerecho){
            this.borDerecho = borDerecho;
            return this;
        }      public Builder setBordeIzquierda(BorderStyle borIzquierda){
            this.borIzquierdo = borIzquierda;
            return this;
        }

        public XSSFCellStyle build(XSSFWorkbook libro){
            XSSFCellStyle estiloCelda = libro.createCellStyle();
            if(this.colorDefecto !=0){
                estiloCelda.setFillForegroundColor(colorDefecto);
            }
            if(this.colorPerzonalizado !=null){
                estiloCelda.setFillForegroundColor(colorPerzonalizado);
            }
            if(this.tipoPatron !=null){
                estiloCelda.setFillPattern(tipoPatron);
            }
            if(this.fuente !=null){
                estiloCelda.setFont(fuente);
            }
            if(this.formato !=null){
                estiloCelda.setDataFormat(libro.createDataFormat().getFormat(formato));
            }
            if(this.alineHorizontal !=null){
                estiloCelda.setAlignment(alineHorizontal);
            }
            if(this.alineVertical !=null){
                estiloCelda.setVerticalAlignment(alineVertical);
            }
            if(this.borArriba !=null){
                estiloCelda.setBorderTop(borArriba);
            }
            if(this.borAbajo !=null){
                estiloCelda.setBorderBottom(borAbajo);
            }
            if(this.borDerecho !=null){
                estiloCelda.setBorderRight(borDerecho);
            }
            if(this.borIzquierdo !=null){
                estiloCelda.setBorderLeft(borIzquierdo);
            }

            return  estiloCelda;

        }




    }
}
