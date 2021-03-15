//import com.google.gson.JsonObject;
//import com.google.gson.JsonParser;
import com.sun.net.httpserver.HttpServer;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.telegram.telegrambots.meta.TelegramBotsApi;
//import org.telegram.telegrambots.meta.exceptions.TelegramApiException;
//import org.telegram.telegrambots.updatesreceivers.DefaultBotSession;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.DocumentFilter;
import javax.swing.text.PlainDocument;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Timer;
import java.util.TimerTask;
import java.util.prefs.Preferences;


public class MainForm extends TrayFrame {
    private JPanel panel1;
    private JButton button1;
    private JTextField textField1;
    private JSpinner spinner1;
    private JTextField textField2;
    private JButton chooseButton;
    public static OfficeShareHttpServer webServer = new OfficeShareHttpServer();
    public static Integer WebPort;
    public static XSSFSheet myExcelSheet;
    public ExelDocInList exelDoc = new ExelDocInList();
    public ArrayList<ExelSheet> exelSheets = new ArrayList<>();


    public MainForm() throws IOException {
        super();
        setContentPane(panel1);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setBounds(250, 250, 800, 350);
        setLocationRelativeTo(null);
        panel1.setBackground(Color.gray);
        panel1.setForeground(Color.white);
        Font font = new Font("Verdana", Font.BOLD, 18);
        panel1.setFont(font);
        setTitle("Office Share");
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setVisible(true);

        button1.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    LoadXLS();
                } catch (IOException ioException) {
                    ioException.printStackTrace();
                }
            }
        });
    }

    public void LoadXLS() throws IOException {
        exelSheets.clear();
        File file = new File("C:\\Users\\id-05\\Dropbox\\PROGRAMS_JAVA\\OfficeShare\\share\\testdoc.xlsx");
        exelDoc.setName(file.getName());
        XSSFWorkbook myExcelBook = null;
        myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        Integer k = 0;
        for(k=0;k<myExcelBook.getNumberOfSheets();k++)
        {
            ArrayList<ArrayList<String>> tableList = new ArrayList<>();
            myExcelSheet = myExcelBook.getSheetAt(k);
            Integer j = 0;
            XSSFRow row = myExcelSheet.getRow(j);
            while ((myExcelSheet.getRow(j) != null) & (!row.getCell(0).toString().equals(""))) {
                Integer i = 0;
                ArrayList<String> bufList = new ArrayList<>();
                while (row.getCell(i) != null) {
                    bufList.add(row.getCell(i).toString());
                    i++;
                }
                tableList.add(bufList);
                j++;
                if (myExcelSheet.getRow(j) != null) {
                    row = myExcelSheet.getRow(j);
                }
            }
            ExelSheet exelSheet = new ExelSheet();
            exelSheet.setName(myExcelSheet.getSheetName());
            exelSheet.setTable(tableList);
            exelSheets.add(exelSheet);
        }
        exelDoc.setSheets(exelSheets);
        myExcelBook.close();
        WebServerStart(exelDoc);
    }

    public static void WebServerStart(ExelDocInList exelDoc){
        WebPort = 99;
        webServer = new OfficeShareHttpServer();
        try {
            webServer.main(WebPort,exelDoc);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        final MainForm MainFormNew = new MainForm();
    }

    private void createUIComponents() {
        // TODO: place custom component creation code here
    }

}
