package RoboSimJava;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
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
    private static final List<String> namesSheetExcel = new ArrayList<>();
    public static ArrayList<String> namesColumnsExcel = new ArrayList<>();
    private static String nameOpenList;
    private static String directoryOpenFile;
    private static ExcelGrouper.ExcelData information;
    private static int titleNumber = 0;
    private static int debitNumber = 1;
    private static int creditNumber = 2;
    private static int dateNumber = 3;

    private ProgressDialog progressDialog;

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
        setSize(800, 600);
        getContentPane().setBackground(new Color(190, 166, 166));
        addComponents();
    }

    private void addComponents() {

        namesSheetExcel.addFirst("выберите лист");
        namesColumnsExcel.addFirst("Выберите столбец");

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
        constraints.gridy = 0;
        constraints.insets = insets;
        panel.add(useListLabel, constraints);

        JComboBox<String> useListField = new JComboBox<>(namesSheetExcel.toArray(new String[0]));
        useListField.addActionListener(e -> {
            nameOpenList = Objects.toString(useListField.getSelectedItem(), "");
        });
        useListField.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 1;
        constraints.gridy = 0;
        constraints.insets = insets;
        panel.add(useListField, constraints);

        JLabel nameListLabel = new JLabel("Название листа:");
        nameListLabel.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 2;
        constraints.gridy = 0;
        constraints.insets = insets;
        panel.add(nameListLabel, constraints);

        JTextField nameListField = new JTextField(14);
        nameListField.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 3;
        constraints.gridy = 0;
        constraints.insets = insets;
        panel.add(nameListField, constraints);

        JLabel titleLabel = new JLabel("Название:");
        titleLabel.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 0;
        constraints.gridy = 2;
        constraints.insets = insets;
        panel.add(titleLabel, constraints);

        JLabel debitLabel = new JLabel("По Дебету:");
        debitLabel.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 1;
        constraints.gridy = 2;
        constraints.insets = insets;
        panel.add(debitLabel, constraints);

        JLabel creditLabel = new JLabel("По Кредиту:");
        creditLabel.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 2;
        constraints.gridy = 2;
        constraints.insets = insets;
        panel.add(creditLabel, constraints);

        JLabel dateLabel = new JLabel("Дата:");
        dateLabel.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 3;
        constraints.gridy = 2;
        constraints.insets = insets;
        panel.add(dateLabel, constraints);

        JComboBox<String> titleBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        titleBox.addActionListener(e -> {
            titleNumber = titleBox.getSelectedIndex() - 1;
        });
        titleBox.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 0;
        constraints.gridy = 3;
        constraints.insets = insets;
        panel.add(titleBox, constraints);

        JComboBox<String> debitBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        debitBox.addActionListener(e -> {
            debitNumber = debitBox.getSelectedIndex() - 1;
        });
        debitBox.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 1;
        constraints.gridy = 3;
        constraints.insets = insets;
        panel.add(debitBox, constraints);

        JComboBox<String> creditBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        creditBox.addActionListener(e -> {
            creditNumber = creditBox.getSelectedIndex() - 1;
        });
        creditBox.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 2;
        constraints.gridy = 3;
        constraints.insets = insets;
        panel.add(creditBox, constraints);

        JComboBox<String> dateBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        dateBox.addActionListener(e -> {
            dateNumber = dateBox.getSelectedIndex() - 1;
        });
        dateBox.setFont(new Font(Font.MONOSPACED, Font.ITALIC, 14));
        constraints.gridx = 3;
        constraints.gridy = 3;
        constraints.insets = insets;
        panel.add(dateBox, constraints);

        JButton generate = new JButton("сгенерировать");
        generate.setPressedIcon(null);
        generate.setContentAreaFilled(false);
        generate.repaint();
        generate.setSize(150, 50);
        constraints.gridx = 3;
        constraints.gridy = 4;
        constraints.insets = insets;
        generate.setFont(fButton);
        panel.add(generate, constraints);

        JButton openExcel = new JButton("открыть");
        openExcel.setPressedIcon(null);
        openExcel.setContentAreaFilled(false);
        openExcel.repaint();
        openExcel.setSize(150, 50);
        constraints.gridx = 0;
        constraints.gridy = 1;
        constraints.insets = insets;
        openExcel.setFont(fButton);
        panel.add(openExcel, constraints);

        JButton readFile = new JButton("читать");
        readFile.setPressedIcon(null);
        readFile.setContentAreaFilled(false);
        readFile.repaint();
        readFile.setSize(150, 50);
        constraints.gridx = 1;
        constraints.gridy = 1;
        constraints.insets = insets;
        readFile.setFont(fButton);
        panel.add(readFile, constraints);

        JButton clearFile = new JButton("очистить");
        clearFile.setPressedIcon(null);
        clearFile.setContentAreaFilled(false);
        clearFile.repaint();
        clearFile.setSize(150, 50);
        constraints.gridx = 2;
        constraints.gridy = 1;
        constraints.insets = insets;
        clearFile.setFont(fButton);
        panel.add(clearFile, constraints);

        JButton saveFile = new JButton("сохранить");
        saveFile.setPressedIcon(null);
        saveFile.setContentAreaFilled(false);
        saveFile.repaint();
        saveFile.setSize(150, 50);
        constraints.gridx = 3;
        constraints.gridy = 1;
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

        DefaultTableModel tableModel = new DefaultTableModel() {
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

        actionButtons(openExcel, useListField, readFile, textArea, tableModel, excelTable, clearFile, saveFile, nameListField, titleBox, debitBox, creditBox, dateBox, generate);

        add(panel);
    }

    private void actionButtons(JButton openExcel, JComboBox<String> useListField, JButton readFile, JTextArea textArea,
                               DefaultTableModel tableModel, JTable excelTable, JButton clearFile, JButton saveFile,
                               JTextField nameListField, JComboBox<String> title, JComboBox<String> debit,
                               JComboBox<String> credit, JComboBox<String> date, JButton generate) {

        openExcel.addActionListener(e -> {
            openExcelDirectory(useListField);
        });

        readFile.addActionListener(e -> {
            readFileWithProgress(textArea, tableModel, excelTable, title, debit, credit, date);
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

        generate.addActionListener(e -> {
            try {
                ExcelGrouper.readSelectedSheet(directoryOpenFile, information,
                        titleNumber, debitNumber, creditNumber, dateNumber);
                ExcelGrouper.cleanData(information);
                JOptionPane.showMessageDialog(this, "Отчет успешно сгенерирован!", "Успех", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException ex) {
                logger.severe("Ошибка при генерации: " + ex.getMessage());
                JOptionPane.showMessageDialog(this, "Ошибка при генерации: " + ex.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
            }
        });
    }

    private void openExcelDirectory(JComboBox<String> useListField) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            logger.severe("Произошла ошибка: открытия окна" + ex.getMessage());
        }

        JFileChooser chooser = FunctionComponent.getFileChooser();
        if (chooser.showDialog(this, "Открыть") == JFileChooser.APPROVE_OPTION) {
            File selectedFile = chooser.getSelectedFile();
            directoryOpenFile = selectedFile.getAbsolutePath();

            progressDialog = new ProgressDialog(this, "Открытие файла");
            progressDialog.setIndeterminate(true);
            progressDialog.setStatus("Анализ файла: " + selectedFile.getName());

            Thread openThread = new Thread(() -> {
                try {
                    information = ExcelGrouper.selectFileAndSheet(directoryOpenFile);

                    SwingUtilities.invokeLater(() -> {
                        namesSheetExcel.clear();
                        changeBox(useListField, FunctionExcel.readSheet(directoryOpenFile), "выберите лист");
                        JOptionPane.showMessageDialog(CreateWindow.this,
                                selectedFile.getName(), "Файл открыт", JOptionPane.INFORMATION_MESSAGE);
                    });
                } catch (IOException e) {
                    logger.severe("Ошибка при открытии файла: " + e.getMessage());
                    SwingUtilities.invokeLater(() -> {
                        JOptionPane.showMessageDialog(CreateWindow.this,
                                "Ошибка при открытии файла: " + e.getMessage(),
                                "Ошибка", JOptionPane.ERROR_MESSAGE);
                    });
                } finally {
                    SwingUtilities.invokeLater(() -> {
                        if (progressDialog != null) {
                            progressDialog.dispose();
                        }
                    });
                }
            });

            openThread.start();
            progressDialog.setVisible(true);
        }
    }

    private void readFileWithProgress(JTextArea textArea, DefaultTableModel tableModel, JTable excelTable,
                                      JComboBox<String> title, JComboBox<String> debit,
                                      JComboBox<String> credit, JComboBox<String> date) {

        if (directoryOpenFile == null) {
            JOptionPane.showMessageDialog(this, "Сначала откройте файл!",
                    "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        if (nameOpenList == null || nameOpenList.isEmpty() || nameOpenList.equals("выберите лист")) {
            JOptionPane.showMessageDialog(this, "Сначала выберите лист!",
                    "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Отключаем кнопки во время чтения
        setButtonsEnabled(false);

        progressDialog = new ProgressDialog(this, "Чтение файла");
        progressDialog.setStatus("Подготовка к чтению...");

        Thread readThread = new Thread(() -> {
            try {
                // Очищаем старые данные в UI потоке
                SwingUtilities.invokeLater(() -> {
                    data.clear();
                    textArea.setText("");
                    tableModel.setRowCount(0);
                    tableModel.setColumnCount(0);
                    namesColumnsExcel.clear();
                    namesColumnsExcel.add("Выберите столбец");
                });

                // Чтение с прогрессом
                FunctionExcel.readWithProgress(nameOpenList, directoryOpenFile, data, progressDialog);

                if (progressDialog.isCancelled()) {
                    SwingUtilities.invokeLater(() -> {
                        JOptionPane.showMessageDialog(CreateWindow.this,
                                "Операция чтения отменена", "Отмена", JOptionPane.INFORMATION_MESSAGE);
                    });
                    return;
                }

                // Отображаем данные в UI потоке
                SwingUtilities.invokeLater(() -> {
                    if (!data.isEmpty()) {
                        FunctionComponent.displayDataInTable(data, tableModel, excelTable);

                        // Показываем в текстовой области только первые 200 строк для производительности
                        int maxDisplay = Math.min(data.size(), 200);
                        for (int i = 1; i <= maxDisplay; i++) {
                            FunctionComponent.appendText(textArea, Arrays.toString(data.get("" + i)) + '\n');
                        }
                        if (data.size() > 200) {
                            FunctionComponent.appendText(textArea, "\n... и еще " + (data.size() - 200) + " строк");
                        }
                    }

                    changeBox(debit, namesColumnsExcel, "Выберите столбец");
                    changeBox(credit, namesColumnsExcel, "Выберите столбец");
                    changeBox(date, namesColumnsExcel, "Выберите столбец");
                    changeBox(title, namesColumnsExcel, "Выберите столбец");

                    JOptionPane.showMessageDialog(CreateWindow.this,
                            "Загружено строк: " + data.size(), "Чтение завершено", JOptionPane.INFORMATION_MESSAGE);
                });

            } catch (Exception ex) {
                logger.severe("Ошибка при чтении файла: " + ex.getMessage());
                SwingUtilities.invokeLater(() -> {
                    JOptionPane.showMessageDialog(CreateWindow.this,
                            "Ошибка при чтении файла: " + ex.getMessage(),
                            "Ошибка", JOptionPane.ERROR_MESSAGE);
                });
            } finally {
                SwingUtilities.invokeLater(() -> {
                    if (progressDialog != null) {
                        progressDialog.dispose();
                    }
                    setButtonsEnabled(true);
                });
            }
        });

        readThread.start();
        progressDialog.setVisible(true);
    }

    private void saveFileDirectory(String nameList) {
        if (information == null || information.rows.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Сначала сгенерируйте отчет!",
                    "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            logger.severe("Произошла ошибка: открытия окна" + ex.getMessage());
        }

        JFileChooser chooser = FunctionComponent.getFileChooser();
        if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            String filePath = chooser.getSelectedFile().getAbsolutePath();
            if (!filePath.toLowerCase().endsWith(".xlsx")) {
                filePath += ".xlsx";
            }

            // Отключаем кнопки во время сохранения
            setButtonsEnabled(false);

            progressDialog = new ProgressDialog(this, "Сохранение файла");
            progressDialog.setIndeterminate(true);
            progressDialog.setStatus("Сохранение отчета...");

            Thread saveThread = getThread(nameList, filePath, chooser);

            saveThread.start();
            progressDialog.setVisible(true);
        }
    }

    private Thread getThread(String nameList, String filePath, JFileChooser chooser) {
        String finalFilePath = filePath;
        Thread saveThread = new Thread(() -> {
            try {

                ExcelGrouper.saveGroupedFileWithYears(finalFilePath, information, "По дебету", "По кредиту", debitNumber, creditNumber);
                FunctionExcel.saveDateInExcel(nameList, finalFilePath, data);

                SwingUtilities.invokeLater(() -> {
                    JOptionPane.showMessageDialog(CreateWindow.this,
                            "Файл успешно сохранен: " + chooser.getSelectedFile().getName(),
                            "Сохранение завершено", JOptionPane.INFORMATION_MESSAGE);
                });

            } catch (IOException e) {
                logger.severe("Ошибка при сохранении: " + e.getMessage());
                SwingUtilities.invokeLater(() -> {
                    JOptionPane.showMessageDialog(CreateWindow.this,
                            "Ошибка при сохранении: " + e.getMessage(),
                            "Ошибка", JOptionPane.ERROR_MESSAGE);
                });
            } finally {
                SwingUtilities.invokeLater(() -> {
                    if (progressDialog != null) {
                        progressDialog.dispose();
                    }
                    setButtonsEnabled(true);
                });
            }
        });
        return saveThread;
    }

    private void setButtonsEnabled(boolean enabled) {
        // Находим все кнопки на форме и включаем/отключаем их
        Component[] components = getContentPane().getComponents();
        for (Component comp : components) {
            if (comp instanceof JPanel) {
                enableComponentsInPanel((JPanel) comp, enabled);
            }
        }
    }

    private void enableComponentsInPanel(JPanel panel, boolean enabled) {
        for (Component comp : panel.getComponents()) {
            if (comp instanceof JButton) {
                comp.setEnabled(enabled);
            } else if (comp instanceof JPanel) {
                enableComponentsInPanel((JPanel) comp, enabled);
            }
        }
    }

    private static void changeBox(JComboBox<String> box, List<String> list, String text) {
        SwingUtilities.invokeLater(() -> {
            box.removeAllItems();
            List<String> tempList = new ArrayList<>(list);
            if (!tempList.contains(text) && !tempList.isEmpty()) {
                tempList.addFirst(text);
            } else if (tempList.isEmpty()) {
                tempList.add(text);
            }

            for (String s : tempList) {
                box.addItem(s);
            }
            box.setSelectedIndex(0);
        });
    }
}