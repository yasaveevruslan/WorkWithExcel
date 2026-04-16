package RoboSimJava;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
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
    private static final Map<String, Object[]> data = new LinkedHashMap<>();
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
    }

    private DefaultTableModel tableModel;

    private void addComponents() {

        namesSheetExcel.addFirst("выберите лист");

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
        openExcel.setPressedIcon(null);
        openExcel.setContentAreaFilled(false);
        openExcel.repaint();
        openExcel.setSize(150, 50);
        constraints.gridx = 0;
        constraints.gridy = 3;
        constraints.insets = insets;
        openExcel.setFont(fButton);
        panel.add(openExcel, constraints);

        JButton readFile = new JButton("читать");
        readFile.setPressedIcon(null);
        readFile.setContentAreaFilled(false);
        readFile.repaint();
        readFile.setSize(150, 50);
        constraints.gridx = 1;
        constraints.gridy = 3;
        constraints.insets = insets;
        readFile.setFont(fButton);
        panel.add(readFile, constraints);

        JButton clearFile = new JButton("очистить");
        clearFile.setPressedIcon(null);
        clearFile.setContentAreaFilled(false);
        clearFile.repaint();
        clearFile.setSize(150, 50);
        constraints.gridx = 2;
        constraints.gridy = 3;
        constraints.insets = insets;
        clearFile.setFont(fButton);
        panel.add(clearFile, constraints);

        JButton saveFile = new JButton("сохранить");
        saveFile.setPressedIcon(null);
        saveFile.setContentAreaFilled(false);
        saveFile.repaint();
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
        constraints.gridy = 5;
        constraints.gridwidth = 4;
        constraints.gridheight = 4;
        constraints.insets = insets;

        JTextArea textArea = new JTextArea();
        textArea.setWrapStyleWord(true);
        textArea.setFont(new Font("Monospaced", Font.PLAIN, 16));

        tableModel = new DefaultTableModel() {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column != 0;
            }
        };
        JTable excelTable = new JTable(tableModel);
        excelTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        excelTable.setFont(new Font("Arial", Font.PLAIN, 14));
        excelTable.setRowHeight(25);
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.addTab("Таблица", new AlwaysScrollableScrollPane(excelTable));
        tabbedPane.addTab("Текст", new AlwaysScrollableScrollPane(textArea));
        mainPanel.add(tabbedPane, BorderLayout.CENTER);
        panel.add(mainPanel, constraints);


        actionButtons(openExcel, useListField, readFile, textArea, tableModel, excelTable, clearFile, saveFile, nameListField);

        add(panel);
    }

    private void actionButtons(JButton openExcel, JComboBox<String> useListField, JButton readFile, JTextArea textArea, DefaultTableModel tableModel, JTable excelTable, JButton clearFile, JButton saveFile, JTextField nameListField) {
        openExcel.addActionListener(e -> {
            openExcelDirectory();
            namesSheetExcel.clear();
            namesSheetExcel = FunctionExcel.readSheet(directoryOpenFile);
            namesSheetExcel.addFirst("выберите лист");
            useListField.removeAllItems();
            for (String s : namesSheetExcel) {
                useListField.addItem(s);
            }
            useListField.setSelectedIndex(0);
        });


        readFile.addActionListener(e -> {
            try {
                FunctionExcel.read(nameOpenList, directoryOpenFile, data);
                textArea.setText("");
                for (int i = 1; i < data.size() + 1; i++) {
                    FunctionComponent.displayDataInTable(data, tableModel, excelTable);
                    FunctionComponent.appendText(textArea, Arrays.toString(data.get("" + i)) + '\n');
                }
                analyzeWithSheetSelection();
            } catch (IOException ex) {
                logger.severe("Произошла ошибка: открытия файла" + ex.getMessage());
                fileHandler.publish(new java.util.logging.LogRecord(Level.SEVERE, "Произошла ошибка: " + ex.getMessage()));
            } catch (InvalidFormatException ex) {
                logger.severe("Произошла ошибка: чтения файла" + ex.getMessage());
                fileHandler.publish(new java.util.logging.LogRecord(Level.SEVERE, "Произошла ошибка: " + ex.getMessage()));
            }
        });

        clearFile.addActionListener(e -> {
            data.clear();
            textArea.setText("");
            tableModel.setRowCount(0);
            tableModel.setColumnCount(0);
            excelTable.revalidate();
            excelTable.repaint();
        });

        saveFile.addActionListener(e -> {
            saveFileDirectory(nameListField.getText());
        });
    }


    private void openExcelDirectory() {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            logger.severe("Произошла ошибка: открытия окна" + ex.getMessage());
            fileHandler.publish(new java.util.logging.LogRecord(Level.SEVERE, "Произошла ошибка: " + ex.getMessage()));
        }

        final JFrame frame = new JFrame("Выбрать");
        JFileChooser chooser = FunctionComponent.getFileChooser();
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
        JFileChooser chooser = FunctionComponent.getFileChooser();
        if (chooser.showDialog(frame, "Сохранить") == JFileChooser.APPROVE_OPTION) {
            String filePath = chooser.getSelectedFile().getAbsolutePath();
            if (!filePath.toLowerCase().endsWith(".xlsx")) {
                filePath += ".xlsx";
            }
            FunctionExcel.saveDateInExcel(nameList, filePath, data);
            JOptionPane.showMessageDialog(null, chooser.getSelectedFile().getName(), "Файл сохранен", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private void analyzeWithSheetSelection() {
        JFileChooser chooser = new JFileChooser();

        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();

            try {
                // МЕТОД 1: Получаем информацию о листах
                ExcelGrouper.ExcelData data = ExcelGrouper.selectFileAndSheet(file.getAbsolutePath());

                // Диалог выбора листа
                String[] sheets = data.sheetNames.toArray(new String[0]);
                String selectedSheet = (String) JOptionPane.showInputDialog(
                        this,
                        "Выберите лист для обработки:",
                        "Выбор листа",
                        JOptionPane.QUESTION_MESSAGE,
                        null,
                        sheets,
                        sheets[0]
                );

                if (selectedSheet != null) {
                    ExcelGrouper.selectSheetByName(data, selectedSheet);

                    // Диалог настройки колонок
                    JPanel panel = new JPanel(new GridLayout(4, 2, 5, 5));
                    JTextField groupColField = new JTextField("0");
                    JTextField val1ColField = new JTextField("1");
                    JTextField val2ColField = new JTextField("2");
                    JTextField val3ColField = new JTextField("3");

                    panel.add(new JLabel("Колонка с названиями:"));
                    panel.add(groupColField);
                    panel.add(new JLabel("Первый числовой столбец:"));
                    panel.add(val1ColField);
                    panel.add(new JLabel("Второй числовой столбец:"));
                    panel.add(val2ColField);
                    panel.add(new JLabel("Дата столбец:"));
                    panel.add(val3ColField);

                    int result = JOptionPane.showConfirmDialog(this, panel,
                            "Настройка колонок", JOptionPane.OK_CANCEL_OPTION);

                    if (result == JOptionPane.OK_OPTION) {
                        int groupCol = Integer.parseInt(groupColField.getText());
                        int val1Col = Integer.parseInt(val1ColField.getText());
                        int val2Col = Integer.parseInt(val2ColField.getText());
                        int val3Col = Integer.parseInt(val3ColField.getText());

                        // МЕТОД 2: Чтение
                        ExcelGrouper.readSelectedSheet(file.getAbsolutePath(), data,
                                groupCol, val1Col, val2Col, val3Col);

                        // МЕТОД 3: Очистка
                        ExcelGrouper.cleanData(data);

                        // Сохранение
                        JFileChooser saveChooser = new JFileChooser();
                        saveChooser.setSelectedFile(new File("grouped_result.xlsx"));

                        if (saveChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
                            String outputPath = saveChooser.getSelectedFile().getAbsolutePath();
                            if (!outputPath.endsWith(".xlsx")) outputPath += ".xlsx";

                            // МЕТОД 4: Сохранение
                            ExcelGrouper.saveGroupedFileWithYears(outputPath, data, "По дебету", "По кредиту", val1Col, val2Col);

                            JOptionPane.showMessageDialog(this, "Файл сохранен!");
                        }
                    }
                }

            } catch (IOException e) {
                JOptionPane.showMessageDialog(this, "Ошибка: " + e.getMessage());
            }
        }
    }
}
