package Controller;

import Browser.AppGui;
import Modell.ExportToExcel;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

public class ControllerExportToExcel implements ActionListener {
    private ExportToExcel modell;
    private AppGui view;
    JFileChooser selecArchivo = new JFileChooser();
    File archivo;
    int contadorAccion = 0;
    private String nam_,surn_,sex_;
    private String[] datos_;

    //Constructor
    public ControllerExportToExcel(AppGui view, ExportToExcel modell) {
        this.view = view;
        this.modell = modell;
    }
    //Method to fix the choice of the format for the export documents
    public void AgregarFiltro() {
        selecArchivo.setFileFilter(new FileNameExtensionFilter("Excel (*xls)", "xls"));
        selecArchivo.setFileFilter(new FileNameExtensionFilter("Excel (*xlsx)", "xlsx"));
    }
    //Action of the events
    public void actionPerformed(ActionEvent e) {
        contadorAccion++;
        if (contadorAccion == 1) {
            AgregarFiltro();
        }
        //Export
        if (e.getSource() == view.exportToExcelButton) {
            if (selecArchivo.showDialog(null, "Export to Excel") == JFileChooser.APPROVE_OPTION) {
                archivo = selecArchivo.getSelectedFile();
                if (archivo.getName().endsWith("xls") || archivo.getName().endsWith("xlsx")) {
                    JOptionPane.showMessageDialog(null, this.modell.Export(archivo, view.table1));
                } else {
                    JOptionPane.showMessageDialog(null, "Select once valid format");
                }
            }
        }  //fin
        //Add die variable to the table
        if (e.getSource() == view.addToListButton) {
            nam_ = view.name.getText(); //take the information
            surn_ = view.surname.getText();
            sex_ = view.sexBox.getSelectedItem().toString();
            String datos_ [] = {nam_,surn_,sex_}; //convert this to a list
            view.tableModel.addRow(datos_); //set the list in the model for show this
        }  //fin
    }
}
