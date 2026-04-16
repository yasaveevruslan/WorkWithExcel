package RoboSimJava;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.*;
import java.util.List;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

public class CreateWindow extends JFrame {

    private static final Logger logger = Logger.getLogger(CreateWindow.class.getName());
    public static FileHandler fileHandler;
    private static Map<String, Object[]> data = new LinkedHashMap<>();
    private static List<String> namesSheetExcel = new ArrayList<String>();
    private static String nameOpenList;

    private static String directoryOpenFile;
    public CreateWindow() {
        initializeWindow();
        setVisible(true);
        try {
            fileHandler = new FileHandler("log.log");
            fileHandler.setFormatter(new SimpleFormatter());
            logger.addHandler(fileHandler);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void initializeWindow() {
        setTitle("Работа с excel");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(700, 600);
        getContentPane().setBackground(new Color(190, 166, 166));
        addComponents();
//        pack();
    }

    private void addComponents() {

        namesSheetExcel.add(0, "выберите лист");

        JPanel panel = new JPanel();
        panel.setLayout(new GridBagLayout());
        GridBagConstraints constraints = new GridBagConstraints();
        Insets insets = new Insets(5, 5, 5, 5);
        constraints.fill = GridBagConstraints.HORIZONTAL;
        constraints.weightx = 1.0;

        Font fButton = new Font(Font.MONOSPACED, Font.BOLD, 18);

        JLabel useListLabel = new JLabel("Выберите лист:");
        useListLabel.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 0;
        constraints.gridy = 2;
        constraints.insets = insets;
        panel.add(useListLabel, constraints);

        JComboBox<String> useListField = new JComboBox<>(namesSheetExcel.toArray(new String[0]));
        useListField.addActionListener(e -> {
        nameOpenList = Objects.toString(useListField.getSelectedItem(), "");
        });
        useListField.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 1;
        constraints.gridy = 2;
        constraints.insets = insets;
        panel.add(useListField, constraints);


        JLabel nameListLabel = new JLabel("Название листа:");
        nameListLabel.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 2;
        constraints.gridy = 2;
        constraints.insets = insets;
        panel.add(nameListLabel, constraints);

        JTextField nameListField = new JTextField(14);
        nameListField.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 3;
        constraints.gridy = 2;
        constraints.insets = insets;
        panel.add(nameListField, constraints);


        JButton openExcel = new JButton("открыть");
        openExcel.setSize(150, 50);
        constraints.gridx = 0;
        constraints.gridy = 3;
        constraints.insets = insets;
        openExcel.setFont(fButton);
        panel.add(openExcel, constraints);

        JButton readFile = new JButton("читать");
        readFile.setSize(150, 50);
        constraints.gridx = 1;
        constraints.gridy = 3;
        constraints.insets = insets;
        readFile.setFont(fButton);
        panel.add(readFile, constraints);

        JButton clearFile = new JButton("очистить");
        clearFile.setSize(150, 50);
        constraints.gridx = 2;
        constraints.gridy = 3;
        constraints.insets = insets;
        clearFile.setFont(fButton);
        panel.add(clearFile, constraints);

        JButton saveFile = new JButton("сохранить");
        saveFile.setSize(150, 50);
        constraints.gridx = 3;
        constraints.gridy = 3;
        constraints.insets = insets;
        saveFile.setFont(fButton);
        panel.add(saveFile, constraints);


        JPanel mainPanel = new JPanel(new BorderLayout());
        constraints.fill = GridBagConstraints.BOTH;
        constraints.weightx = 1.0;
        constraints.weighty = 1.0;
        constraints.gridx = 0;
        constraints.gridy = 4;
        constraints.gridwidth = 4;
        constraints.gridheight = 4;
        constraints.insets = insets;

        JTextArea textArea = new JTextArea();
        textArea.setWrapStyleWord(true);
        textArea.setFont(new Font("Monospaced", Font.PLAIN, 16));

        DefaultTableModel tableModel = new DefaultTableModel();
        JTable excelTable = new JTable(tableModel);
        excelTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        excelTable.setFont(new Font("Arial", Font.PLAIN, 14));
        excelTable.setRowHeight(25);
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.addTab("Таблица", new AlwaysScrollableScrollPane(excelTable));
        tabbedPane.addTab("Текст", new AlwaysScrollableScrollPane(textArea));
        mainPanel.add(tabbedPane, BorderLayout.CENTER);
        panel.add(mainPanel, constraints);



        openExcel.addActionListener(e -> {
            openExcelDirectory();
            namesSheetExcel.clear();
            namesSheetExcel = FunctionExcel.readSheet(directoryOpenFile);
            useListField.removeAllItems();
            for (String s : namesSheetExcel) {
                useListField.addItem(s);
            }
            useListField.setSelectedIndex(0);
        });

        readFile.addActionListener(e -> {
            try {
//                FunctionExcel.read(useListField.getText(), directoryOpenFile, data);
                FunctionExcel.read(nameOpenList, directoryOpenFile, data);
                for (int i = 1; i < data.size() + 1; i++) {
                    displayDataInTable(data, tableModel, excelTable);        // Обновляем таблицу
                    appendText(textArea, Arrays.toString(data.get("" + i)) + '\n');
                }

            } catch (IOException ex) {
                logger.severe("Произошла ошибка: открытия файла" + ex.getMessage());
                fileHandler.publish(new java.util.logging.LogRecord(Level.SEVERE, "Произошла ошибка: " + ex.getMessage()));
            } catch (InvalidFormatException ex) {
                logger.severe("Произошла ошибка: чтения файла" + ex.getMessage());
                fileHandler.publish(new java.util.logging.LogRecord(Level.SEVERE, "Произошла ошибка: " + ex.getMessage()));
            }
        });

        saveFile.addActionListener(e -> {
            saveFileDirectory(nameListField.getText());
        });

        add(panel);
    }


    private void openExcelDirectory() {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            logger.severe("Произошла ошибка: открытия окна" + ex.getMessage());
            fileHandler.publish(new java.util.logging.LogRecord(Level.SEVERE, "Произошла ошибка: " + ex.getMessage()));
        }

        final JFrame frame = new JFrame("Выбрать");
        JFileChooser chooser = getFileChooser();
        if (chooser.showDialog(frame, "Открыть") == JFileChooser.APPROVE_OPTION) {
            directoryOpenFile = chooser.getSelectedFile().getAbsolutePath();
            JOptionPane.showMessageDialog(null, chooser.getSelectedFile().getName(), "Название файла", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private void saveFileDirectory(String nameList) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            logger.severe("Произошла ошибка: открытия окна" + ex.getMessage());
            fileHandler.publish(new java.util.logging.LogRecord(Level.SEVERE, "Произошла ошибка: " + ex.getMessage()));
        }

        final JFrame frame = new JFrame("Сохранить");
        JFileChooser chooser = getFileChooser();
        if (chooser.showDialog(frame, "Сохранить") == JFileChooser.APPROVE_OPTION) {
            String filePath = chooser.getSelectedFile().getAbsolutePath();
            if (!filePath.toLowerCase().endsWith(".xlsx")) {
                filePath += ".xlsx";
            }
            FunctionExcel.saveDateInExcel(nameList, filePath, data);
            JOptionPane.showMessageDialog(null, chooser.getSelectedFile().getName(), "Файл сохранен", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private static JFileChooser getFileChooser() {
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



    private void displayDataInTable(Map<String, Object[]> data, DefaultTableModel tableModel, JTable excelTable) {
        tableModel.setRowCount(0);
        tableModel.setColumnCount(0);

        if (data.isEmpty()) {
            return;
        }

        int maxCols = 0;
        for (Object[] row : data.values()) {
            maxCols = Math.max(maxCols, row.length);
        }

        String[] columns = new String[maxCols];
        for (int i = 0; i < maxCols; i++) {
            columns[i] = getColumnLetter(i);
        }
        tableModel.setColumnIdentifiers(columns);

        for (String key : data.keySet()) {
            Object[] row = data.get(key);
            Object[] tableRow = new Object[maxCols];

            for (int i = 0; i < maxCols; i++) {
                if (i < row.length && row[i] != null) {
                    tableRow[i] = row[i].toString();
                } else {
                    tableRow[i] = "";
                }
            }

            tableModel.addRow(tableRow);
        }

        for (int i = 0; i < maxCols; i++) {
            excelTable.getColumnModel().getColumn(i).setPreferredWidth(100);
        }
    }

    private String getColumnLetter(int index) {
        StringBuilder sb = new StringBuilder();
        while (index >= 0) {
            sb.insert(0, (char) ('A' + index % 26));
            index = index / 26 - 1;
        }
        return sb.toString();
    }
}
