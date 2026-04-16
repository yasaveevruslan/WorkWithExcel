package RoboSimJava;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.util.Map;

public class FunctionComponent {

    public static JFileChooser getFileChooser() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            @Override
            public boolean accept(File f) {
                if (f.isDirectory()) return true;
                String name = f.getName().toLowerCase();
                return name.endsWith(".xlsx") || name.endsWith(".xls");
            }

            @Override
            public String getDescription() {
                return "Excel Files (*.xlsx, *.xls)";
            }
        });
        chooser.setAcceptAllFileFilterUsed(false);
        return chooser;
    }


    public static void appendText(JTextArea textArea, String text) {
        textArea.append(text);
        textArea.setCaretPosition(textArea.getDocument().getLength());
    }



    public static void displayDataInTable(Map<String, Object[]> data, DefaultTableModel tableModel, JTable excelTable) {

        tableModel.setRowCount(0);
        tableModel.setColumnCount(0);

        if (data.isEmpty()) {
            return;
        }

        int maxCols = 0;
        for (Object[] row : data.values()) {
            maxCols = Math.max(maxCols, row.length);
        }

        String[] columns = new String[maxCols + 1];
        columns[0] = "№";
        for (int i = 0; i < maxCols; i++) {
            columns[i + 1] = getColumnLetter(i);
        }
        tableModel.setColumnIdentifiers(columns);

        int rowNumber = 1;
        for (String key : data.keySet()) {
            Object[] row = data.get(key);
            Object[] tableRow = new Object[maxCols + 1];

            tableRow[0] = rowNumber++;

            for (int i = 0; i < maxCols; i++) {
                if (i < row.length && row[i] != null) {
                    tableRow[i + 1] = row[i].toString();
                } else {
                    tableRow[i + 1] = "";
                }
            }

            tableModel.addRow(tableRow);
        }

        excelTable.getColumnModel().getColumn(0).setPreferredWidth(50);
        excelTable.getColumnModel().getColumn(0).setMaxWidth(60);
        excelTable.getColumnModel().getColumn(0).setMinWidth(40);



        for (int i = 1; i < maxCols; i++) {
            excelTable.getColumnModel().getColumn(i).setPreferredWidth(100);
        }



        DefaultTableCellRenderer rowHeaderRenderer = new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value,
                                                           boolean isSelected, boolean hasFocus,
                                                           int row, int column) {

                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

                c.setBackground(new Color(240, 240, 240)); // Светло-серый фон
                c.setFont(c.getFont().deriveFont(Font.BOLD)); // Жирный шрифт
                setHorizontalAlignment(CENTER); // Выравнивание по центру

                setBorder(BorderFactory.createMatteBorder(0, 0, 1, 1, Color.GRAY));
                setEnabled(false);
                return c;
            }
        };

        excelTable.getColumnModel().getColumn(0).setCellRenderer(rowHeaderRenderer);
        excelTable.getColumnModel().getColumn(0).setCellEditor(null);
    }

    public static String getColumnLetter(int index) {
        StringBuilder sb = new StringBuilder();
        while (index >= 0) {
            sb.insert(0, (char) ('A' + index % 26));
            index = index / 26 - 1;
        }
        return sb.toString();
    }
}
