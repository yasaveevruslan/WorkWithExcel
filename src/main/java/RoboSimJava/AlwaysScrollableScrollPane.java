package RoboSimJava;

import javax.swing.*;
import javax.swing.event.AncestorEvent;
import javax.swing.event.AncestorListener;
import java.awt.*;

public class AlwaysScrollableScrollPane extends JScrollPane {

    public AlwaysScrollableScrollPane(Component view) {
        super(view);

        // Сразу настраиваем прокрутку
        setWheelScrollingEnabled(true);
        getVerticalScrollBar().setUnitIncrement(20);
        getHorizontalScrollBar().setUnitIncrement(20);

        // Добавляем слушатель для ранней инициализации
        addHierarchyListener(e -> {
            if (isShowing()) {
                requestFocusInWindow();
            }
        });

        // Принудительная инициализация при добавлении
        addAncestorListener(new AncestorListener() {
            @Override
            public void ancestorAdded(AncestorEvent event) {
                SwingUtilities.invokeLater(() -> {
                    setWheelScrollingEnabled(true);
                    requestFocusInWindow();
                });
            }

            @Override
            public void ancestorRemoved(AncestorEvent event) {}

            @Override
            public void ancestorMoved(AncestorEvent event) {}
        });
    }

    @Override
    public void addNotify() {
        super.addNotify();
        // При добавлении в иерархию компонентов
        SwingUtilities.invokeLater(() -> {
            setWheelScrollingEnabled(true);
            getVerticalScrollBar().setUnitIncrement(20);
        });
    }
}
