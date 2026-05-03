package RoboSimJava;

import javax.swing.*;
import java.awt.*;

public class ProgressDialog extends JDialog {
    private final JProgressBar progressBar;
    private final JLabel statusLabel;
    private volatile boolean cancelled = false;

    public ProgressDialog(JFrame parent, String title) {
        super(parent, title, true);
        setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);

        setLayout(new BorderLayout());
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));

        statusLabel = new JLabel(" ", SwingConstants.CENTER);
        progressBar = new JProgressBar(0, 100);
        progressBar.setStringPainted(true);

        JButton cancelButton = new JButton("Отмена");
        cancelButton.addActionListener(e -> {
            cancelled = true;
            statusLabel.setText("Отмена...");
            cancelButton.setEnabled(false);
        });

        JPanel buttonPanel = new JPanel();
        buttonPanel.add(cancelButton);

        panel.add(statusLabel, BorderLayout.NORTH);
        panel.add(progressBar, BorderLayout.CENTER);
        panel.add(buttonPanel, BorderLayout.SOUTH);

        add(panel);
        setSize(400, 150);
        setLocationRelativeTo(parent);
    }

    public void setProgress(int progress) {
        SwingUtilities.invokeLater(() -> {
            progressBar.setValue(progress);
            progressBar.setString(progress + "%");
        });
    }

    public void setStatus(String status) {
        SwingUtilities.invokeLater(() -> statusLabel.setText(status));
    }

    public void setIndeterminate(boolean indeterminate) {
        SwingUtilities.invokeLater(() -> {
            progressBar.setIndeterminate(indeterminate);
            if (indeterminate) {
                progressBar.setString("Подождите...");
            }
        });
    }

    public boolean isCancelled() {
        return cancelled;
    }
}