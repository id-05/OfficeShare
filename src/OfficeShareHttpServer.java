import com.sun.net.httpserver.*;
import java.io.IOException;
import java.io.OutputStream;
import java.net.InetSocketAddress;
import java.util.ArrayList;

public class OfficeShareHttpServer {
    HttpServer server;
    ExelDocInList exelDoc;

    public void main(int WebPort, ExelDocInList exelDoc) throws Exception {
        this.exelDoc = exelDoc;
        server = HttpServer.create();
        server.bind(new InetSocketAddress(WebPort), 0);
        HttpContext context = server.createContext("/", new EchoHandler());
        context.setAuthenticator(new Auth());
        server.setExecutor(null);
        server.start();
    }

    class EchoHandler implements HttpHandler {
        @Override
        public void handle(HttpExchange exchange) throws IOException {
            StringBuilder builder = new StringBuilder();
            builder.append("<head>");
            builder.append("<meta charset=\"utf-8\">");
            builder.append("<title>Office Share</title>");
            builder.append("<style type=\"text/css\">");
            builder.append("</style>");
            builder.append("</head>");
            builder.append("<body bgcolor = \"#ffffff\" text = \"#000000\">");
            builder.append("<p style=\"text-align: center;\">&nbsp;</p>");
            builder.append("<h1 style=\"text-align: center;\"><strong>"+exelDoc.getName()+"</strong></h1>");
            builder.append("<p style=\"text-align: center;\">&nbsp;</p>");
            builder.append("<p>&nbsp;</p>");


            for(ExelSheet bufSheet:exelDoc.getSheets()) {
                ArrayList<ArrayList<String>> tableList = bufSheet.getTable();
                builder.append("<h2>" + bufSheet.getName() + "</h2>");
                builder.append("<table style=\"height: 68px; border-color: black; width: 800px; margin-left: auto; margin-right: auto;\" border=\"4\" cellpadding=\"4\"><caption>&nbsp;</caption>");
                builder.append("<tbody>");
                builder.append("<tr>");
                for (int i = 0; i < tableList.size(); i++) {
                    builder.append("<h1><tr>");
                    ArrayList<String> bufList = tableList.get(i);
                    for (int j = 0; j < bufList.size(); j++) {
                        builder.append("<td style=\"width: 41px; text-align: center;\">" + bufList.get(j) + "</td>");
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
            byte[] bytes = builder.toString().getBytes();
            exchange.sendResponseHeaders(200, bytes.length);
            OutputStream os = exchange.getResponseBody();
            os.write(bytes);
            os.close();
        }
    }

    static class Auth extends Authenticator {
        @Override
        public Result authenticate(HttpExchange httpExchange) {
            if ("/json".equals(httpExchange.getRequestURI().toString()))
                return new Success(new HttpPrincipal("c0nst", "realm"));
            else
                return new Success(new HttpPrincipal("c0nst", "realm"));
        }
    }
}
