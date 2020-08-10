package Browser;

import Controller.ControllerExportToExcel;
import Modell.ExportToExcel;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.table.DefaultTableModel;
import java.awt.*;



public class AppGui extends JFrame {
    public JTextField name;
    public JTextField surname;
    public JComboBox sexBox;
    public JTable table1;
    public JButton exportToExcelButton;
    private JButton addToDBButton;
    private JPanel panelAppGui;
    public JButton addToListButton;
    public DefaultTableModel tableModel;


    public AppGui() {
        //Design
        panelAppGui = new JPanel();
        panelAppGui.setLayout(new GridBagLayout());
        panelAppGui.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Amorettyss", TitledBorder.DEFAULT_JUSTIFICATION, TitledBorder.DEFAULT_POSITION, null, null));
        final JLabel label1 = new JLabel();
        label1.setText("Type on the field the name, surname and Sexo ");
        GridBagConstraints gbc;
        gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.anchor = GridBagConstraints.WEST;
        panelAppGui.add(label1, gbc);
        //******Edited*********
        //*Table and Container*
        String[][] NOM = {};
        String[] colName = {"Name", "Surname", "Sexo"};
        //Crating of Container for scrolling, create and add one model to table. Edding die Table to container
        tableModel = new DefaultTableModel(NOM, colName);
        table1 = new JTable(); //new JTable (NOM, colName);
        table1.setModel(tableModel);
        //create scroll to show value for the table
        table1.setPreferredScrollableViewportSize(new Dimension(300, 70));
        table1.setFillsViewportHeight(true);
        //Create the scroll pane and add the table to it
        JScrollPane scrollPane = new JScrollPane(table1);
        gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 10;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.fill = GridBagConstraints.BOTH;
        panelAppGui.add(scrollPane, gbc);
        /*table1.setVisible(true);
        gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 10;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.fill = GridBagConstraints.BOTH;
        panelAppGui.add(table1, gbc);
         */
        //End Edited
        name = new JTextField();
        name.setText("Name");
        gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 6;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.anchor = GridBagConstraints.WEST;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panelAppGui.add(name, gbc);
        surname = new JTextField();
        surname.setText("Surname");
        gbc = new GridBagConstraints();
        gbc.gridx = 1;
        gbc.gridy = 6;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.anchor = GridBagConstraints.WEST;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panelAppGui.add(surname, gbc);
        addToListButton = new JButton();
        addToListButton.setText("Add to list people");
        gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 7;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panelAppGui.add(addToListButton, gbc);
        sexBox = new JComboBox();
        final DefaultComboBoxModel defaultComboBoxModel1 = new DefaultComboBoxModel();
        defaultComboBoxModel1.addElement("M");
        defaultComboBoxModel1.addElement("F");
        defaultComboBoxModel1.addElement("Other");
        sexBox.setModel(defaultComboBoxModel1);
        gbc = new GridBagConstraints();
        gbc.gridx = 2;
        gbc.gridy = 6;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.anchor = GridBagConstraints.WEST;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panelAppGui.add(sexBox, gbc);
        exportToExcelButton = new JButton();
        exportToExcelButton.setText("Export to Excel");
        gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 14;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panelAppGui.add(exportToExcelButton, gbc);
        addToDBButton = new JButton();
        addToDBButton.setText("Add to DB");
        gbc = new GridBagConstraints();
        gbc.gridx = 1;
        gbc.gridy = 14;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panelAppGui.add(addToDBButton, gbc);

        //Call to Button actions
        addToListButton.addActionListener(new ControllerExportToExcel(this, new ExportToExcel()));

        exportToExcelButton.addActionListener(new ControllerExportToExcel(this, new ExportToExcel()));

    }

    public static void main(String[] args) {

        JFrame frame = new JFrame("AppGui");
        frame.setContentPane(new AppGui().panelAppGui);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);

    }

}
