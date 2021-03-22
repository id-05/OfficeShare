import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.FilterChain;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.swing.*;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;

@WebServlet("/hello")
public class MainServlet extends HttpServlet {

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        String FilePath;
        XSSFSheet myExcelSheet;
        ExelDocInList exelDoc = new ExelDocInList();
        ArrayList<ExelSheet> exelSheets = new ArrayList<>();

        File file = new File("D:/testdoc.xlsx");
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

        StringBuilder builder = new StringBuilder();
        builder.append("<head>");
        builder.append("<meta charset=\"utf-8\">");
        builder.append("<title>Office Share</title>");
        builder.append("<style type=\"text/css\">");
        builder.append("</style>");
        builder.append("</head>");
        builder.append("<body bgcolor = \"#ffffff\" text = \"#000000\">");
        builder.append("<p style=\"text-align: center;\">&nbsp;</p>");
        builder.append("<h1 style=\"text-align: center;\"><strong>" + exelDoc.getName() + "</strong></h1>");
        builder.append("<p style=\"text-align: center;\">&nbsp;</p>");
        builder.append("<p>&nbsp;</p>");
        for (ExelSheet bufSheet : exelDoc.getSheets()) {
            builder.append("&nbsp;&nbsp;<a href=#" + bufSheet.getName() + ">&nbsp;" + bufSheet.getName() + "&nbsp;</a>&nbsp;&nbsp;&nbsp;");
        }
        builder.append("<p>&nbsp;</p>");
        builder.append("&nbsp;&nbsp;<a href=/download?id="+"testdoc.xlsx"+">&nbsp;" +  "ТЕСТ ССЫЛКИ" + "</a>;");
        builder.append("<h2 style=\"text-align: center;\"><input type=\"button\" onclick='window.location.reload()' value=\"RELOAD\" /></h2>");
        builder.append("<p>&nbsp;</p>");
        for (ExelSheet bufSheet : exelDoc.getSheets()) {
            ArrayList<ArrayList<String>> tableList = bufSheet.getTable();
            builder.append("&nbsp;&nbsp;<p>&nbsp;&nbsp;<a name=" + bufSheet.getName() + "></a></p>");
            builder.append("<h2>" + bufSheet.getName() + "</h2>");
            builder.append("<table style=\"height: 68px; border-color: black; width: 800px; margin-left: auto; margin-right: auto;\" border=\"4\" cellpadding=\"4\"><caption>&nbsp;</caption>");
            builder.append("<tbody>");
            builder.append("<tr>");
            for (int i = 0; i < tableList.size(); i++) {
                builder.append("<h1><tr>");
                ArrayList<String> bufList = tableList.get(i);
                for (int j = 0; j < bufList.size(); j++) {
                    if(j==10){
                        builder.append("<td style=\"width: 41px; text-align: center;\">" +
                                "<a href=/download?id="+ URLEncoder.encode((bufList.get(j)),"UTF-8")+">" +  bufList.get(j) + "</a>"
                                + "</td>");
                    }else {
                        builder.append("<td style=\"width: 41px; text-align: center;\">" + bufList.get(j) + "</td>");
                    }
                }
                builder.append("</tr></h1>");

            }
            builder.append("</tbody>");
            builder.append("</table>");
            builder.append("<p>&nbsp;</p>");
        }

        builder.append("<p>&nbsp;</p>");
        builder.append("<p>&nbsp;</p>");
        builder.append("<p>&nbsp;</p>");
        builder.append("<p>&nbsp;</p>");
        builder.append("<p>&nbsp;</p>");
        builder.append("<p>&nbsp;</p>");
        builder.append("<h2 style=\"text-align: center;\"><input type=\"button\" onclick='window.location.reload()' value=\"RELOAD\" /></h2>");
        builder.append("</body>");
        builder.append("</html>");
        req.setCharacterEncoding("UTF-8");
        resp.setCharacterEncoding("UTF-8");
        resp.setContentType("text/html");
        PrintWriter printWriter = resp.getWriter();
        printWriter.write(String.valueOf(builder));
        printWriter.close();
    }
}

