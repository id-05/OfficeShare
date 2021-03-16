import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.IOException;
import java.net.URL;

public class TrayFrame extends JFrame {
    public static TrayFrame app;
    private final TrayIcon trayIcon;
    private final SystemTray systemTray = SystemTray.getSystemTray();
    public boolean chetTray = false;

    public TrayFrame() throws IOException {
        super("Minimize to tray");
        URL imageURL = this.getClass().getResource("/Ikonka.png");
        Image icon = Toolkit.getDefaultToolkit().getImage(imageURL);
        trayIcon = new TrayIcon(icon, "Minimize to tray");
        trayIcon.setImageAutoSize(true);
        trayIcon.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent ev) {
                setVisible(true);
                setState(JFrame.NORMAL);
                removeTray();
            }
        });

        MouseListener mouS = new MouseListener() {
            public void mouseClicked(MouseEvent ev) {
            }

            public void mouseEntered(MouseEvent ev) {
            }

            public void mouseExited(MouseEvent ev) {
            }

            public void mousePressed(MouseEvent ev) {
            }

            public void mouseReleased(MouseEvent ev) {
            }
        };
        trayIcon.addMouseListener(mouS);
        MouseMotionListener mouM = new MouseMotionListener() {
            public void mouseDragged(MouseEvent ev) {
            }

            public void mouseMoved(MouseEvent ev) {
                trayIcon.setToolTip("Double click - expand");
            }
        };

        trayIcon.addMouseMotionListener(mouM);
        addWindowStateListener(ev -> {
            if (ev.getNewState() == JFrame.ICONIFIED) {
                setVisible(false);
                addTray();
            }
        });
    }


    private void removeTray() {
        systemTray.remove(trayIcon);
    }

    private void addTray() {
        try {
            systemTray.add(trayIcon);
            chetTray = true;
        } catch (AWTException ex) {
            ex.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        app = new TrayFrame();
        app.setVisible(true);
        app.setAlwaysOnTop(true);

        app.addWindowListener(new WindowListener() {
            public void windowClosing(WindowEvent winEvent) {
                System.exit(0);
            }

            public void windowActivated(WindowEvent winEvent) {
            }

            public void windowClosed(WindowEvent winEvent) {
            }

            public void windowDeactivated(WindowEvent winEvent) {
            }

            public void windowDeiconified(WindowEvent winEvent) {
            }

            public void windowIconified(WindowEvent winEvent) {
            }

            public void windowOpened(WindowEvent winEvent) {
            }
        });
    }
}

