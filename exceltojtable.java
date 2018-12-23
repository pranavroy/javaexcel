/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exceltojtable;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.*;
//import javafx.scene.control.Cell;
//import javafx.scene.control.Cell;
import javax.swing.table.DefaultTableModel;
import jxl.*;
import jxl.read.biff.BiffException;
public class exceltojtable extends JFrame {

    static JTable table;
    static JScrollPane scroll;
    // header is Vector contains table Column
    static Vector headers = new Vector();
    // Model is used to construct JTable
    static DefaultTableModel model = null;
    // data is Vector contains Data from Excel File
    static Vector data = new Vector();
    static JButton jbClick,jbinsert;
    static JFileChooser jChooser;
    static int tableWidth = 0; // set the tableWidth
    static int tableHeight = 0; // set the tableHeight

    public exceltojtable() {
        super("Import Excel To JTable");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JPanel buttonPanel = new JPanel();
        buttonPanel.setBackground(Color.white);
        jChooser = new JFileChooser();
        jbClick = new JButton("Select Excel File");
        jbinsert=new JButton("Insert data");
        buttonPanel.add(jbinsert,BorderLayout.EAST);
        buttonPanel.add(jbClick, BorderLayout.CENTER);
        // Show Button Click Event
        jbClick.addActionListener(new ActionListener() {

            public void actionPerformed(ActionEvent arg0) {
                jChooser.showOpenDialog(null);

                File file = jChooser.getSelectedFile();
                if (!file.getName().endsWith("xls")) {
                    JOptionPane.showMessageDialog(null,
                            "Please select only Excel file.",
                            "Error", JOptionPane.ERROR_MESSAGE);
                } else {
                    try {
                        fillData(file);
                    } catch (IOException ex) {
                        Logger.getLogger(exceltojtable.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    model = new DefaultTableModel(data,
                            headers);
                    tableWidth = model.getColumnCount()
                            * 150;
                    tableHeight = model.getRowCount()
                            * 25;
                    table.setPreferredSize(new Dimension(
                            tableWidth, tableHeight));

                    table.setModel(model);
                }
            }
        });
        jbClick.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            }
        });

        table = new JTable();
        table.setAutoCreateRowSorter(true);
        
        model = new DefaultTableModel(data, headers);

        table.setModel(model);
        table.setBackground(Color.pink);

        table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        table.setEnabled(false);
        table.setRowHeight(25);
        table.setRowMargin(4);

        tableWidth = model.getColumnCount() * 150;
        tableHeight = model.getRowCount() * 25;
        table.setPreferredSize(new Dimension(
                tableWidth, tableHeight));

        scroll = new JScrollPane(table);
        scroll.setBackground(Color.pink);
        scroll.setPreferredSize(new Dimension(300, 300));
        scroll.setHorizontalScrollBarPolicy(
                JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        scroll.setVerticalScrollBarPolicy(
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        getContentPane().add(buttonPanel,
                BorderLayout.NORTH);
        getContentPane().add(scroll,
                BorderLayout.CENTER);
        GraphicsDevice gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
        int width = gd.getDisplayMode().getWidth();
        int height = gd.getDisplayMode().getHeight();
        setSize(width, height);
        setResizable(true);
        setVisible(true);
    }

    /**
     * Fill JTable with Excel file data.
     *
     * @param file file :contains xls file to display in jTable
     */
    void fillData(File file) throws IOException {

        Workbook workbook = null;
        try {
            workbook = Workbook.getWorkbook(file);
            Sheet sheet = workbook.getSheet(0);

            headers.clear();
            for (int i = 0; i < sheet.getColumns(); i++) {
                Cell cell1 = sheet.getCell(i, 0);
                headers.add(cell1.getContents());
            }

            data.clear();
            for (int j = 1; j < sheet.getRows(); j++) {
                Vector d = new Vector();
                for (int i = 0; i < sheet.getColumns(); i++) {

                    Cell cell = sheet.getCell(i, j);

                    d.add(cell.getContents());

                }
                d.add("\n");
                data.add(d);
            }
        } catch (BiffException e) {
            e.printStackTrace();
        }
        
    }

    public static void main(String[] args) {

        new exceltojtable();
    }
}
