package RoboSimJava;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import javax.swing.border.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.*;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.logging.FileHandler;
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
    private static int titleNumber = -1;
    private static int debitNumber = -1;
    private static int creditNumber = -1;
    private static int dateNumber = -1;
    private static int innNumber = -1;
    private static int operationNameNumber = -1;

    private ProgressDialog progressDialog;
    private JComboBox<String> useListField;
    private JComboBox<String> titleBox, debitBox, creditBox, dateBox, innBox, operationBox;
    private JTextArea textArea;
    private DefaultTableModel tableModel;
    private JTable excelTable;
    private JTextArea infoArea;
    private JTextField nameListField;

    // Цветовая схема
    private static final Color COLOR_BACKGROUND = new Color(245, 245, 250);
    private static final Color COLOR_PANEL = new Color(255, 255, 255);
    private static final Color COLOR_PRIMARY = new Color(41, 128, 185);
    private static final Color COLOR_SUCCESS = new Color(39, 174, 96);
    private static final Color COLOR_WARNING = new Color(243, 156, 18);
    private static final Color COLOR_DANGER = new Color(231, 76, 60);
    private static final Color COLOR_BORDER = new Color(200, 200, 210);
    private static final Color COLOR_TEXT = new Color(44, 62, 80);
    private static final Color COLOR_LABEL = new Color(52, 73, 94);

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
        setTitle("Финансовый отчет");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1200, 800);
        setLocationRelativeTo(null);
        getContentPane().setBackground(COLOR_BACKGROUND);

        ImageIcon appIcon = loadApplicationIcon();
        if (appIcon != null) {
            setIconImage(appIcon.getImage());
        } else {
            // Если иконка не загрузилась, используем запасной вариант
            setIconImage(createFallbackIcon());
        }

        addComponents();

    }

    private ImageIcon loadApplicationIcon() {
        try {

            // Способ 1
            URL imageUrl = getClass().getClassLoader().getResource("app-icon.png");
            if (imageUrl != null) {
                System.out.println("Иконка найдена: " + imageUrl);
                return new ImageIcon(imageUrl);
            }

            // Способ 2
            imageUrl = getClass().getResource("/app-icon.png");
            if (imageUrl != null) {
                System.out.println("Иконка найдена: " + imageUrl);
                return new ImageIcon(imageUrl);
            }

            // Способ 3
            String[] iconNames = {"icon.png", "logo.png", "app.png", "icon.ico"};
            for (String name : iconNames) {
                imageUrl = getClass().getClassLoader().getResource(name);
                if (imageUrl != null) {
                    System.out.println("Иконка найдена: " + name);
                    return new ImageIcon(imageUrl);
                }
            }

            System.out.println("Иконка не найдена в ресурсах. Будет использована стандартная иконка.");
            return null;

        } catch (Exception e) {
            System.err.println("Ошибка при загрузке иконки: " + e.getMessage());
            return null;
        }
    }

    private Image createFallbackIcon() {
        java.awt.image.BufferedImage image = new java.awt.image.BufferedImage(32, 32, java.awt.image.BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = image.createGraphics();

        // Включаем сглаживание
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

        // Рисуем фон с градиентом
        GradientPaint gradient = new GradientPaint(0, 0, COLOR_PRIMARY, 32, 32, new Color(31, 97, 141));
        g2d.setPaint(gradient);
        g2d.fillRoundRect(2, 2, 28, 28, 8, 8);

        // Белая окантовка
        g2d.setColor(Color.WHITE);
        g2d.setStroke(new BasicStroke(1.5f));
        g2d.drawRoundRect(2, 2, 28, 28, 8, 8);

        // Рисуем символ таблицы
        g2d.setColor(Color.WHITE);
        g2d.setFont(new Font("Segoe UI", Font.BOLD, 18));
        FontMetrics fm = g2d.getFontMetrics();
        String symbol = "";
        int x = (32 - fm.stringWidth(symbol)) / 2;
        int y = ((32 - fm.getHeight()) / 2) + fm.getAscent();
        g2d.drawString(symbol, x, y);

        g2d.dispose();
        return image;
    }

    private void addComponents() {
        namesSheetExcel.addFirst("выберите лист");
        namesColumnsExcel.addFirst("Выберите столбец");

        setLayout(new BorderLayout(10, 10));

        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBackground(COLOR_BACKGROUND);
        mainPanel.setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        mainPanel.add(createSettingsPanel(), BorderLayout.NORTH);
        mainPanel.add(createDataPanel(), BorderLayout.CENTER);
        mainPanel.add(createInfoPanel(), BorderLayout.SOUTH);

        add(mainPanel);
    }

    private JPanel createSettingsPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBackground(COLOR_PANEL);
        panel.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(COLOR_BORDER, 1, true),
                BorderFactory.createEmptyBorder(15, 20, 15, 20)
        ));

        JLabel titleLabel = new JLabel("Настройки отчета");
        titleLabel.setFont(new Font("Segoe UI", Font.BOLD, 18));
        titleLabel.setForeground(COLOR_PRIMARY);
        titleLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.add(titleLabel);
        panel.add(Box.createRigidArea(new Dimension(0, 15)));

        JPanel filePanel = createFilePanel();
        filePanel.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.add(filePanel);
        panel.add(Box.createRigidArea(new Dimension(0, 15)));

        JPanel columnsPanel = createColumnsPanel();
        columnsPanel.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.add(columnsPanel);
        panel.add(Box.createRigidArea(new Dimension(0, 15)));

        JPanel actionPanel = createActionPanel();
        actionPanel.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.add(actionPanel);

        return panel;
    }

    private JPanel createFilePanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBackground(COLOR_PANEL);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        Font labelFont = new Font("Segoe UI", Font.PLAIN, 13);
        Font fieldFont = new Font("Segoe UI", Font.PLAIN, 13);

        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0;
        JLabel sheetLabel = new JLabel("Лист:");
        sheetLabel.setFont(labelFont);
        sheetLabel.setForeground(COLOR_LABEL);
        panel.add(sheetLabel, gbc);

        gbc.gridx = 1;
        gbc.weightx = 0.3;
        useListField = new JComboBox<>(namesSheetExcel.toArray(new String[0]));
        useListField.addActionListener(e -> {
            nameOpenList = Objects.toString(useListField.getSelectedItem(), "");
        });
        useListField.setFont(fieldFont);
        panel.add(useListField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        JLabel nameLabel = new JLabel("Лист сохранения:");
        nameLabel.setFont(labelFont);
        nameLabel.setForeground(COLOR_LABEL);
        panel.add(nameLabel, gbc);

        gbc.gridx = 3;
        gbc.weightx = 0.4;
        nameListField = new JTextField(20);
        nameListField.setFont(fieldFont);
        nameListField.setBorder(createRoundedBorder());
        panel.add(nameListField, gbc);

        return panel;
    }

    private JPanel createColumnsPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBackground(COLOR_PANEL);

        JLabel mappingLabel = new JLabel("Соответствие столбцов");
        mappingLabel.setFont(new Font("Segoe UI", Font.BOLD, 14));
        mappingLabel.setForeground(COLOR_LABEL);
        mappingLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.add(mappingLabel);
        panel.add(Box.createRigidArea(new Dimension(0, 10)));

        JPanel columnsGrid = new JPanel(new GridLayout(2, 3, 15, 10));
        columnsGrid.setBackground(COLOR_PANEL);

        Font labelFont = new Font("Segoe UI", Font.PLAIN, 12);
        Font comboFont = new Font("Segoe UI", Font.PLAIN, 12);

        String[] columnTypes = {"Наименование:", "Дебет:", "Кредит:", "Дата:", "ИНН:", "Операция:"};

        // Создаем панель для каждого комбобокса
        JPanel[] itemPanels = new JPanel[6];

        titleBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        titleBox.addActionListener(e -> titleNumber = titleBox.getSelectedIndex() - 1);
        titleBox.setFont(comboFont);

        debitBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        debitBox.addActionListener(e -> debitNumber = debitBox.getSelectedIndex() - 1);
        debitBox.setFont(comboFont);

        creditBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        creditBox.addActionListener(e -> creditNumber = creditBox.getSelectedIndex() - 1);
        creditBox.setFont(comboFont);

        dateBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        dateBox.addActionListener(e -> dateNumber = dateBox.getSelectedIndex() - 1);
        dateBox.setFont(comboFont);

        innBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        innBox.addActionListener(e -> innNumber = innBox.getSelectedIndex() - 1);
        innBox.setFont(comboFont);

        operationBox = new JComboBox<>(namesColumnsExcel.toArray(new String[0]));
        operationBox.addActionListener(e -> operationNameNumber = operationBox.getSelectedIndex() - 1);
        operationBox.setFont(comboFont);

        JComboBox<?>[] boxes = {titleBox, debitBox, creditBox, dateBox, innBox, operationBox};

        for (int i = 0; i < columnTypes.length; i++) {
            itemPanels[i] = new JPanel(new BorderLayout(5, 0));
            itemPanels[i].setBackground(COLOR_PANEL);

            JLabel label = new JLabel(columnTypes[i]);
            label.setFont(labelFont);
            label.setForeground(COLOR_LABEL);

            itemPanels[i].add(label, BorderLayout.NORTH);
            itemPanels[i].add(boxes[i], BorderLayout.CENTER);
            columnsGrid.add(itemPanels[i]);
        }

        panel.add(columnsGrid);

        return panel;
    }

    private JPanel createActionPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.CENTER, 15, 0));
        panel.setBackground(COLOR_PANEL);

        Font buttonFont = new Font("Segoe UI", Font.BOLD, 14);

        JButton openExcel = createStyledButton("Открыть файл", COLOR_PRIMARY, buttonFont);
        JButton readFile = createStyledButton("Прочитать данные", COLOR_SUCCESS, buttonFont);
        JButton clearFile = createStyledButton("Очистить", COLOR_WARNING, buttonFont);
        JButton saveFile = createStyledButton("Сохранить отчет", COLOR_SUCCESS, buttonFont);
        JButton generate = createStyledButton("СГЕНЕРИРОВАТЬ ОТЧЕТ", new Color(155, 89, 182), new Font("Segoe UI", Font.BOLD, 16));

        panel.add(openExcel);
        panel.add(readFile);
        panel.add(clearFile);
        panel.add(saveFile);
        panel.add(generate);

        // Добавляем обработчики кнопок
        openExcel.addActionListener(e -> openExcelDirectory());
        readFile.addActionListener(e -> readFileWithProgress());
        clearFile.addActionListener(e -> clearData());
        saveFile.addActionListener(e -> saveFileDirectory());
        generate.addActionListener(e -> generateReport());

        return panel;
    }

    private JPanel createDataPanel() {
        JPanel panel = new JPanel(new BorderLayout(0, 0));
        panel.setBackground(COLOR_PANEL);
        panel.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(COLOR_BORDER, 1, true),
                BorderFactory.createEmptyBorder(10, 10, 10, 10)
        ));

        textArea = new JTextArea();
        textArea.setWrapStyleWord(true);
        textArea.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        textArea.setBackground(new Color(250, 250, 252));

        tableModel = new DefaultTableModel() {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        excelTable = new JTable(tableModel);
        excelTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        excelTable.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        excelTable.setRowHeight(25);
        excelTable.setGridColor(COLOR_BORDER);
        excelTable.setSelectionBackground(new Color(41, 128, 185, 50));

        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        tabbedPane.addTab("Таблица", new AlwaysScrollableScrollPane(excelTable));
        tabbedPane.addTab("Текст", new AlwaysScrollableScrollPane(textArea));

        panel.add(tabbedPane, BorderLayout.CENTER);

        return panel;
    }

    private JPanel createInfoPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(new Color(250, 250, 252));
        panel.setBorder(BorderFactory.createCompoundBorder(
                new LineBorder(COLOR_BORDER, 1, true),
                BorderFactory.createEmptyBorder(10, 15, 10, 15)
        ));

        infoArea = new JTextArea(4, 50);
        infoArea.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        infoArea.setBackground(new Color(250, 250, 252));
        infoArea.setForeground(COLOR_TEXT);
        infoArea.setEditable(false);
        infoArea.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        infoArea.setText(
                "Инструкция:\n" +
                        "1. Откройте Excel файл с данными\n" +
                        "2. Выберите лист и нажмите 'Прочитать данные'\n" +
                        "3. Укажите соответствие столбцов (название, дебет, кредит, дата)\n" +
                        "4. Нажмите 'Сгенерировать отчет' для создания финансового отчета\n" +
                        "5. Сохраните результат через 'Сохранить отчет'"
        );

        JScrollPane infoScroll = new JScrollPane(infoArea);
        infoScroll.setBorder(null);
        infoScroll.setBackground(new Color(250, 250, 252));

        panel.add(infoScroll, BorderLayout.CENTER);

        return panel;
    }

    private JButton createStyledButton(String text, Color bgColor, Font font) {
        JButton button = new JButton(text);
        button.setFont(font);
        button.setBackground(bgColor);
        button.setForeground(Color.WHITE);
        button.setFocusPainted(false);
        button.setBorder(BorderFactory.createEmptyBorder(10, 20, 10, 20));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));

        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                button.setBackground(bgColor.darker());
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                button.setBackground(bgColor);
            }
        });

        return button;
    }

    private Border createRoundedBorder() {
        return BorderFactory.createCompoundBorder(
                new LineBorder(COLOR_BORDER, 1, true),
                BorderFactory.createEmptyBorder(5, 8, 5, 8)
        );
    }

    // ==================== МЕТОДЫ ДЛЯ РАБОТЫ С ДАННЫМИ ====================

    private void openExcelDirectory() {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            logger.severe("Ошибка при открытии окна: " + ex.getMessage());
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
                        infoArea.setText("Файл открыт: " + selectedFile.getName() +
                                "\nКоличество листов: " + information.sheetNames.size() +
                                "\nВыберите лист и нажмите 'Прочитать данные'");
                        JOptionPane.showMessageDialog(CreateWindow.this,
                                selectedFile.getName() + "\nКоличество листов: " + information.sheetNames.size(),
                                "Файл открыт", JOptionPane.INFORMATION_MESSAGE);
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

    private void readFileWithProgress() {
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

        setButtonsEnabled(false);

        progressDialog = new ProgressDialog(this, "Чтение файла");
        progressDialog.setStatus("Подготовка к чтению...");

        Thread readThread = new Thread(() -> {
            try {
                SwingUtilities.invokeLater(() -> {
                    data.clear();
                    textArea.setText("");
                    tableModel.setRowCount(0);
                    tableModel.setColumnCount(0);
                    namesColumnsExcel.clear();
                    namesColumnsExcel.add("Выберите столбец");
                });

                FunctionExcel.readWithProgress(nameOpenList, directoryOpenFile, data, progressDialog);

                if (progressDialog.isCancelled()) {
                    SwingUtilities.invokeLater(() -> {
                        JOptionPane.showMessageDialog(CreateWindow.this,
                                "Операция чтения отменена", "Отмена", JOptionPane.INFORMATION_MESSAGE);
                    });
                    return;
                }

                SwingUtilities.invokeLater(() -> {
                    if (!data.isEmpty()) {
                        FunctionComponent.displayDataInTable(data, tableModel, excelTable);

                        int maxDisplay = Math.min(data.size(), 200);
                        for (int i = 1; i <= maxDisplay; i++) {
                            FunctionComponent.appendText(textArea, Arrays.toString(data.get("" + i)) + '\n');
                        }
                        if (data.size() > 200) {
                            FunctionComponent.appendText(textArea, "\n... и еще " + (data.size() - 200) + " строк");
                        }
                    }

                    changeBox(debitBox, namesColumnsExcel, "Выберите столбец");
                    changeBox(creditBox, namesColumnsExcel, "Выберите столбец");
                    changeBox(dateBox, namesColumnsExcel, "Выберите столбец");
                    changeBox(titleBox, namesColumnsExcel, "Выберите столбец");
                    changeBox(innBox, namesColumnsExcel, "Выберите столбец");
                    changeBox(operationBox, namesColumnsExcel, "Выберите столбец");

                    infoArea.setText("Данные загружены успешно!\n" +
                            "Всего строк: " + data.size() + "\n" +
                            "Всего столбцов: " + (namesColumnsExcel.size() - 1) + "\n" +
                            "Выберите соответствие столбцов и нажмите 'СГЕНЕРИРОВАТЬ ОТЧЕТ'");

                    JOptionPane.showMessageDialog(CreateWindow.this,
                            "Загружено строк: " + data.size() + "\nЗагружено столбцов: " + (namesColumnsExcel.size() - 1),
                            "Чтение завершено", JOptionPane.INFORMATION_MESSAGE);
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

    private void clearData() {
        data.clear();
        textArea.setText("");
        tableModel.setRowCount(0);
        tableModel.setColumnCount(0);
        excelTable.revalidate();
        excelTable.repaint();
        infoArea.setText("Данные очищены");
    }

    private void saveFileDirectory() {
        if (information == null || information.rows.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Сначала сгенерируйте отчет!",
                    "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            logger.severe("Ошибка при открытии окна: " + ex.getMessage());
        }

        JFileChooser chooser = FunctionComponent.getFileChooser();
        if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            String filePath = chooser.getSelectedFile().getAbsolutePath();
            if (!filePath.toLowerCase().endsWith(".xlsx")) {
                filePath += ".xlsx";
            }

            setButtonsEnabled(false);

            progressDialog = new ProgressDialog(this, "Сохранение файла");
            progressDialog.setIndeterminate(true);
            progressDialog.setStatus("Сохранение отчета...");

            String finalFilePath = filePath;
            Thread saveThread = new Thread(() -> {
                try {
                    ExcelGrouper.saveReportWithAllColumns(finalFilePath, information,
                            "Отчет по дебету",
                            "Отчет по кредиту",
                            titleNumber, debitNumber, creditNumber, dateNumber,
                            innNumber, operationNameNumber);

                    String sheetName = nameListField.getText().trim();
                    if (sheetName.isEmpty()) {
                        sheetName = "Исходные данные";
                    }
                    FunctionExcel.saveDateInExcel(sheetName, finalFilePath, data);

                    SwingUtilities.invokeLater(() -> {
                        JOptionPane.showMessageDialog(CreateWindow.this,
                                "Файл успешно сохранен: " + chooser.getSelectedFile().getName(),
                                "Сохранение завершено", JOptionPane.INFORMATION_MESSAGE);
                        infoArea.setText("Отчет сохранен: " + chooser.getSelectedFile().getName());
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

            saveThread.start();
            progressDialog.setVisible(true);
        }
    }

    private void generateReport() {
        if (data == null || data.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Сначала прочитайте данные из файла! (нажмите 'Прочитать данные')",
                    "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        if (titleNumber < 0 || debitNumber < 0 || creditNumber < 0) {
            JOptionPane.showMessageDialog(this, "Выберите все необходимые столбцы!\n" +
                            "Должны быть выбраны: Название, Дебет, Кредит",
                    "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        if (dateNumber < 0) {
            JOptionPane.showMessageDialog(this, "Выберите столбец с датой!\n" +
                            "Это необходимо для фильтрации данных",
                    "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try {
            if (information == null) {
                information = new ExcelGrouper.ExcelData();
            }
            information.rows.clear();

            int skippedNoDate = 0;
            int skippedZero = 0;
            int processed = 0;

            for (Map.Entry<String, Object[]> entry : data.entrySet()) {
                Object[] rowData = entry.getValue();

                String dateString = (dateNumber >= 0 && dateNumber < rowData.length && rowData[dateNumber] != null) ?
                        rowData[dateNumber].toString().trim() : "";

                Date dateValue = null;
                boolean hasValidDate = false;

                if (!dateString.isEmpty()) {
                    dateValue = parseDate(dateString);
                    hasValidDate = (dateValue != null);
                }

                if (!hasValidDate) {
                    skippedNoDate++;
                    continue;
                }

                String groupName = (titleNumber >= 0 && titleNumber < rowData.length && rowData[titleNumber] != null) ?
                        rowData[titleNumber].toString().trim() : "";

                if (groupName.isEmpty()) {
                    skippedNoDate++;
                    continue;
                }

                double value1 = (debitNumber >= 0 && debitNumber < rowData.length) ?
                        getNumericValueFromObject(rowData[debitNumber]) : 0;
                double value2 = (creditNumber >= 0 && creditNumber < rowData.length) ?
                        getNumericValueFromObject(rowData[creditNumber]) : 0;

                if (value1 == 0 && value2 == 0) {
                    skippedZero++;
                    continue;
                }

                ExcelGrouper.DataRow dataRow = new ExcelGrouper.DataRow(groupName, rowData, value1, value2);
                dataRow.setDate(dateValue);

                if (innNumber >= 0 && innNumber < rowData.length && rowData[innNumber] != null) {
                    dataRow.setInn(rowData[innNumber].toString());
                }
                if (operationNameNumber >= 0 && operationNameNumber < rowData.length && rowData[operationNameNumber] != null) {
                    dataRow.setOperationName(rowData[operationNameNumber].toString());
                }

                information.rows.add(dataRow);
                processed++;
            }

            int filteredCount = data.size() - information.rows.size();

            infoArea.setText(String.format(
                    "Отчет сгенерирован!\n" +
                            "Всего строк: %d | Обработано: %d | Отфильтровано: %d\n" +
                            "   • Без даты: %d | • Нулевые суммы: %d",
                    data.size(), information.rows.size(), filteredCount, skippedNoDate, skippedZero
            ));

            if (information.rows.isEmpty()) {
                JOptionPane.showMessageDialog(this,
                        "Не найдено строк с корректными датами!\n" +
                                "Проверьте правильность выбора столбца с датой.",
                        "Предупреждение", JOptionPane.WARNING_MESSAGE);
                return;
            }

            ExcelGrouper.processDataForReport(information, titleNumber, debitNumber, creditNumber,
                    dateNumber, innNumber, operationNameNumber);
            ExcelGrouper.cleanData(information);

            JOptionPane.showMessageDialog(this, String.format(
                            "Отчет успешно сгенерирован!\n\n" +
                                    "Всего строк в файле: %d\n" +
                                    "Обработано записей: %d\n" +
                                    "Отфильтровано строк: %d\n" +
                                    "   • Без корректной даты: %d\n" +
                                    "   • С нулевыми суммами: %d",
                            data.size(), information.rows.size(), filteredCount, skippedNoDate, skippedZero),
                    "Успех", JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception ex) {
            logger.severe("Ошибка при генерации: " + ex.getMessage());
            ex.printStackTrace();
            JOptionPane.showMessageDialog(this, "Ошибка при генерации: " + ex.getMessage(),
                    "Ошибка", JOptionPane.ERROR_MESSAGE);
        }
    }

    // ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================

    private void setButtonsEnabled(boolean enabled) {
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

    private void changeBox(JComboBox<String> box, List<String> list, String text) {
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

    private double getNumericValueFromObject(Object obj) {
        if (obj == null) return 0;
        try {
            if (obj instanceof Number) {
                return ((Number) obj).doubleValue();
            }
            String str = obj.toString().replace(",", ".").replace(" ", "").replace(" ", "");
            if (str.isEmpty()) return 0;
            return Double.parseDouble(str);
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private Date parseDate(String dateStr) {
        if (dateStr == null || dateStr.isEmpty()) return null;

        dateStr = dateStr.trim();

        String[] patterns = {
                "dd.MM.yyyy", "dd.MM.yy", "yyyy-MM-dd", "dd/MM/yyyy",
                "dd/MM/yy", "MM/dd/yyyy", "yyyy/MM/dd", "dd-MM-yyyy",
                "dd-MM-yy", "yyyyMMdd", "ddMMyyyy", "dd.MM.yyyy HH:mm:ss",
                "yyyy-MM-dd HH:mm:ss", "dd.MM.yy HH:mm:ss"
        };

        for (String pattern : patterns) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                sdf.setLenient(false);
                return sdf.parse(dateStr);
            } catch (Exception e) {
                // Пробуем следующий формат
            }
        }

        return null;
    }
}