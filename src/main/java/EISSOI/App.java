package EISSOI;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class App {

    public static FileWriter writer;
    public static Connection con = null;

    // Если собираем проект и тестим на СЕРВЕРЕ то используем URL, PASS, USER такого вида И УБИРАЕМ ЗАВИСИМОСТЬ sourceforge
     public static String url = "jdbc:sqlserver://10.2515.160.75:1433;databaseName=REPORTDATA;integratedSecurity=true";
    // Если собираем проект и тестим на локальной машине I-Novus то используем URL такого вида
/*
        public static String url = "jdbc:jtds:sqlserver://10.255.160.75;databaseName=REPORTDATA;integratedSecurity=true;Domain=GISOMS";
        public static String user = "Apatronov";
        public static String password = "";
*/

    public static void main(String[] args) {
        try {
            writer = new FileWriter("C:/1/java3.txt", false);

            // Если собираем проект и тестим на СЕРВЕРЕ то используем CON такого вида
            System.out.println("try to connect");
            con = DriverManager.getConnection(url);
            System.out.println("after connect");
            // Если собираем проект и тестим на локальной машине I-Novus то используем CON такого вида и подключаем jtds dependency в pom
            //con = DriverManager.getConnection(url, user, password);
            System.out.println("go00");
            File dir = new File("D:/Reports_Outgoing/");
            File[] arrFiles = dir.listFiles();
            System.out.println("go0");
            List<File> lst = Arrays.asList(arrFiles);
            System.out.println(lst.size());
            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
            System.out.println("go");
            // если пытаемся загрузить историю за X дней. готовим переменную dateBefore7Days
            Date d = new Date();
            Calendar cal = Calendar.getInstance();
            System.out.println("go2");
            cal.setTime(d);
            System.out.println("go3");
            writer.append(cal.getTime().toString());
            System.out.println("go4");
            cal.add(Calendar.DATE, -10);
            Date dateBefore7Days = cal.getTime();

            for (int i = 0; i < lst.size(); i++) {
                File file = new File(lst.get(i).toString());
                System.out.println("file");
                final long lastModified = file.lastModified();
                //System.out.println(file.getName() + ",  " + file.lastModified());
                // Если пытаемся загрузить файлики за ИНТЕРВАЛ
                 //if (new Date(lastModified).after(dateBefore7Days)) {

                // Если пытаемся загрузить файлики за текущий день
                if (sdf.format(new Date(lastModified)).equals(sdf.format(d))) {

                    final String filename = file.getAbsolutePath();
                    // извлекаем из имени файла дату
                    StringBuffer file_data = new StringBuffer(filename.substring(filename.length() - 12, filename.length() - 4));
                    String source = file_data.toString();
                    final DateTimeFormatter srcFormatter = DateTimeFormatter.ofPattern("ddMMyyyy");
                    final DateTimeFormatter trgFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd");
                    final LocalDate date = LocalDate.parse(file_data, srcFormatter);
                    final String target = trgFormatter.format(date);
                    //System.out.println(source + " => " + target);

                    Reader reader = null;

                    if (file.getName().toString().substring(0, 3).equals("EMF")) {
                        reader = new EMF_2_reader(filename, target);
                        reader.startread(con);
                    }

                    if (file.getName().toString().substring(0, 3).equals("ПВГ")) {
                        reader = new PVG_5_reader(filename, target);
                        reader.startread(con);
                    }

                    if (file.getName().toString().substring(0, 5).equals("ЧЗЛ_1")) {
                        reader = new CHZL1_4_reader(filename, target);
                        reader.startread(con);
                    }

                    if (file.getName().toString().substring(0, 3).equals("ПМО")) {
                        reader = new PMO_3_reader(filename, target);
                        reader.startread(con);
                    }
                    if (file.getName().toString().substring(0, 4).equals("ЧЗЛ7")) {
                        reader = new CHZL7_1_reader(filename, target);
                        reader.startread(con);
                    }

                    if (file.getName().toString().substring(0, 4).equals("МПНВ")) {
                        reader = new MPNV_6_reader(filename, target);
                        reader.startread(con);
                    }
                    if (file.getName().toString().substring(0, 6).equals("Policy")) {
                        reader = new PolicyTypes_8_reader(filename, target);
                        reader.startread(con);
                    }

                    if (file.getName().toString().substring(0, 3).equals("УЭК")) {
                        reader = new UEK_7_reader(filename, target);
                        reader.startread(con);
                    }

                }
            }
            con.close();
        } catch (
                Exception e) {
            e.getLocalizedMessage();
            System.out.println(e.getMessage());
        }
    }

    public static void connecting(Connection con, String f, String QUERRRY) {
        Statement st = null;
        try {
            st = con.createStatement();
            st.executeQuery(QUERRRY);
            st.close();
        } catch (SQLException ex) {
            if (!ex.getMessage().toString().equals("The statement did not return a result set."))
                System.out.println(ex.getMessage());
            try {
                if (!ex.getMessage().toString().equals("The statement did not return a result set.")) {
                    writer.append('\n');
                    writer.append('\n');
                    writer.append(f);
                    writer.append('\n');
                    writer.append(QUERRRY);
                    writer.append('\n');
                    writer.append(ex.getMessage());
                    writer.append('\n');
                    writer.flush();
                }
            } catch (Exception e) {
            }
        }
    }
}