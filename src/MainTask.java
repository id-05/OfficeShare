import sun.applet.Main;

import javax.swing.*;
import java.io.File;
import java.io.IOException;
import java.lang.management.ManagementFactory;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.TimerTask;

public class MainTask extends TimerTask {

    public MainTask(){

    }

    @Override
    public void run() {
        try {
            MainForm.LoadXLS();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}