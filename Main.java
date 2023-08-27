import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {
    private static String[][] array;
    private static DefaultTableModel model;
    private static JTable table;
    private static JFrame frame;
    private static int selectedRow;

    public static void main(String[] args) {
        array = new String[10][6];
        Integer count_r = 0;
        Integer count_c = 0;

        try {
            File file = new File("C:\\Users\\tsinco-pc10\\Documents\\MyProject\\Source File\\Excel-3692.xlsx");
            FileInputStream fis = new FileInputStream(file);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> itr = sheet.iterator();
            for (int i = 0; i <= 9; i++) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                for (int j = 0; j <= 5; j++) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t\t\t");
                            array[i][j] = cell.getStringCellValue();
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t\t\t");
                            array[i][j] = cell.getStringCellValue();
                            break;
                        default:
                    }
                    count_c++;
                }
                System.out.println("");
                count_r++;
            }

            frame = new JFrame();
            String[] columnNames = {"کدملی", "نام", "نام خانوادگی", "نام پدر", "ناریخ تولد", "شماره شناسنامه"};

            model = new DefaultTableModel(array, columnNames);
            table = new JTable(model);

            JButton add = new JButton("Add");
            JButton edit = new JButton("Edit");
            JButton delete = new JButton("Delete");

            add.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent ae) {
                    openInputWindow(false);
                }
            });

            delete.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent ae) {
                    int selectedRow = table.getSelectedRow();
                    if (selectedRow >= 0) {
                        model.removeRow(selectedRow);
                    } else {
                        JOptionPane.showMessageDialog(frame, "No row selected.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            });

            edit.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent ae) {
                    selectedRow = table.getSelectedRow();
                    if (selectedRow >= 0) {
                        openInputWindow(true);
                    } else {
                        JOptionPane.showMessageDialog(frame, "No row selected.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            });

            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            Container c = frame.getContentPane();
            c.setLayout(new BorderLayout());
            JPanel buttonPanel = new JPanel();
            buttonPanel.add(add);
            buttonPanel.add(delete);
            buttonPanel.add(edit);
            c.add(buttonPanel, BorderLayout.SOUTH);
            c.add(new JScrollPane(table), BorderLayout.CENTER);

            frame.pack();
            frame.setLocationRelativeTo(null);
            frame.setVisible(true);

            add.setBackground(Color.GREEN);
            delete.setBackground(Color.red);

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void openInputWindow(boolean isEdit) {
        JFrame inputFrame = new JFrame();
        inputFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        inputFrame.setLayout(new BorderLayout());

        JPanel inputPanel = new JPanel(new GridLayout(0, 2)); // Use GridLayout to arrange labels and text fields in a grid

        JLabel[] labels = new JLabel[6];
        JTextField[] textFields = new JTextField[6];
        String[] labelNames = {"کدملی", "نام", "نام خانوادگی", "نام پدر", "ناریخ تولد", "شماره شناسنامه"};

        for (int i = 0; i < 6; i++) {
            labels[i] = new JLabel(labelNames[i]);
            textFields[i] = new JTextField(10);
            inputPanel.add(labels[i]);
            inputPanel.add(textFields[i]);

            if (isEdit) {
                String cellValue = (String) table.getValueAt(selectedRow, i);
                textFields[i].setText(cellValue);
            }
        }

        inputFrame.add(inputPanel, BorderLayout.NORTH); // Add input panel to the top of the frame

        JButton addButton = new JButton(isEdit ? "Save" : "Add");

        addButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent ae) {
                List<String> inputs = new ArrayList<>();
                for (int i = 0; i < 6; i++) {
                    String input = textFields[i].getText();
                    inputs.add(input);
                }
                if (!inputs.isEmpty()) {
                    if (isEdit) {
                        model.removeRow(selectedRow);
                        model.insertRow(selectedRow, inputs.toArray());
                    } else {
                        model.addRow(inputs.toArray());
                    }
                    inputFrame.dispose();
                }
            }
        });

        inputFrame.add(addButton, BorderLayout.CENTER); // Add button to the center of the frame

        inputFrame.pack();
        inputFrame.setLocationRelativeTo(null);
        inputFrame.setVisible(true);
    }
}
