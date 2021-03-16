import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
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
    private JButton startButton;
    private JTextField textField1;
    private JSpinner spinner1;
    private JTextField textField2;
    private JButton chooseButton;
    private JButton openWeb;
    private JButton saveButton;
    public static OfficeShareHttpServer webServer = new OfficeShareHttpServer();
    public static Integer WebPort;
    public static String FilePath;
    public static XSSFSheet myExcelSheet;
    public static ExelDocInList exelDoc = new ExelDocInList();
    public static ArrayList<ExelSheet> exelSheets = new ArrayList<>();
    public static Preferences userPrefs;
    public static Integer timeToUpdate;
    public static TimerTask timerTask;
    public static Timer timer;



    public MainForm() throws IOException {
        super();
        userPrefs = Preferences.userRoot().node("OfficeShare");
        setContentPane(panel1);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setBounds(250, 250, 400, 350);
        setLocationRelativeTo(null);
        panel1.setBackground(Color.gray);
        panel1.setForeground(Color.white);
        Font font = new Font("Verdana", Font.BOLD, 18);
        panel1.setFont(font);
        setTitle("Office Share");
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setVisible(true);

        startButton.addActionListener(e -> {
            loadConfig();
            try {
                LoadXLS();
            } catch (IOException ioException) {
                ioException.printStackTrace();
            }
            WebServerStart(exelDoc);
        });
        openWeb.addActionListener(e -> {
            try {
                Desktop.getDesktop().browse(new URI("http://localhost:"+WebPort));
            } catch (IOException | URISyntaxException ex) {
                ex.printStackTrace();
            }
        });

        saveButton.addActionListener(e -> saveConfig());

        chooseButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
            FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX documents", "xlsx");
            fileChooser.addChoosableFileFilter(filter);
            int result = fileChooser.showOpenDialog(MainForm.this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                System.out.println("Selected file: " + selectedFile.getAbsolutePath());
                FilePath = selectedFile.getAbsolutePath();
                textField2.setText(FilePath);
            }
        });

        loadConfig();
        if(!FilePath.equals("empty")){
            LoadXLS();
        }
        WebServerStart(exelDoc);
        timerStart();
    }

    public void loadConfig(){
        WebPort = userPrefs.getInt("WebPort", 99);
        timeToUpdate = userPrefs.getInt("TimeToUpdate", 99);
        FilePath = userPrefs.get("FilePath", "empty");
        textField2.setText(FilePath);
        textField1.setText(WebPort.toString());
        SpinnerModel spinnerModel = new SpinnerNumberModel(1,1,9999,1);//Model[arr];
        spinner1.setModel(spinnerModel);
        spinnerModel.setValue(timeToUpdate);
    }

    public void saveConfig(){
        MainForm.userPrefs.putInt("WebPort",Integer.parseInt(textField1.getText()));
        MainForm.userPrefs.putInt("TimeToUpdate", ((Integer) spinner1.getValue()));
        MainForm.userPrefs.put("FilePath", textField2.getText());
    }

    public static void LoadXLS() throws IOException {
        exelSheets.clear();
        try {
            File file = new File(FilePath);
            exelDoc.setName(file.getName());
            XSSFWorkbook myExcelBook;
            myExcelBook = new XSSFWorkbook(new FileInputStream(file));
            for (int k = 0; k < myExcelBook.getNumberOfSheets(); k++) {
                ArrayList<ArrayList<String>> tableList = new ArrayList<>();
                myExcelSheet = myExcelBook.getSheetAt(k);
                int j = 0;
                XSSFRow row = myExcelSheet.getRow(j);
                while ((myExcelSheet.getRow(j) != null) & (!row.getCell(0).toString().equals(""))) {
                    int i = 0;
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
        }catch (Exception e){
            JOptionPane.showMessageDialog(null,"LOAD FILE FAILED!");
        }
    }

    public static void timerStart(){
        timerTask = null;
        timer = null;
        timerTask = new MainTask();
        timer = new Timer(true);
        timer.scheduleAtFixedRate(timerTask, 0, timeToUpdate *60*1000);
    }

    public static void WebServerStart(ExelDocInList exelDoc){
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

    private static void createUIComponents() {
        JFrame frame = new JFrame("Swing Tester");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        // TODO: place custom component creation code here
    }

}
