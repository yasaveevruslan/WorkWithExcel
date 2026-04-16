package RoboSimJava;

import javax.swing.*;
import javax.swing.event.AncestorEvent;
import javax.swing.event.AncestorListener;
import java.awt.*;

public class AlwaysScrollableScrollPane extends JScrollPane {

    public AlwaysScrollableScrollPane(Component view) {
        super(view);

        setWheelScrollingEnabled(true);
        getVerticalScrollBar().setUnitIncrement(20);
        getHorizontalScrollBar().setUnitIncrement(20);

        addHierarchyListener(e -> {
            if (isShowing()) {
                requestFocusInWindow();
            }
        });

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
        SwingUtilities.invokeLater(() -> {
            setWheelScrollingEnabled(true);
            getVerticalScrollBar().setUnitIncrement(20);
        });
    }
}
