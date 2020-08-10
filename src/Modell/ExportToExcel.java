package Modell;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;

public class ExportToExcel {
Workbook wb;
        public String Export (File archivo, JTable table){
        String respuesta = "It could did not be exported";
        int numeroFila = table.getRowCount();
        int numeroColumna = table.getColumnCount();
        if (archivo.getName().endsWith("xls")){
                wb = new HSSFWorkbook();
        }else{
                wb = new XSSFWorkbook(); //here will be created the document on the program but still not be saved on the PC
        }
        Sheet hoja = wb.createSheet("Book_Test");
        try {
                for (int i = -1; i < numeroFila; i++){ //the row -1 is the row for the title of the table in excel
                        Row fila = hoja.createRow(i+1);
                        for (int j=0; j < numeroColumna; j++){
                                Cell celda = fila.createCell(j);
                                if(i==-1){
                                        celda.setCellValue(String.valueOf(table.getColumnName(j))); //Assignation of Row titles
                                }else {
                                        celda.setCellValue(String.valueOf(table.getValueAt(i, j)));
                                }
                                wb.write(new FileOutputStream(archivo)); //saving the doc on the pc
                        }
                }
                respuesta="Eureka! Exported to Excel";
        } catch(Exception e){
        }
        return respuesta;
        }
}
