package EISSOI;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class App {

    public static FileWriter writer;

    //public static String url = "jdbc:jtds:sqlserver://10.255.160.75;databaseName=REPORTDATA;integratedSecurity=true;Domain=GISOMS";
    // public static String user = "Apatronov";
    //public static String password = "N0vusadm3";
    public static Connection con = null;
    public static String url = "jdbc:sqlserver://10.255.160.75:1433;databaseName=REPORTDATA;integratedSecurity=true";

    public static void main(String[] args) {
        try {
            writer = new FileWriter("C:/1/java3.txt", false);

            //con = DriverManager.getConnection(url, user, password);
            con = DriverManager.getConnection(url);

            File dir = new File("D:/Reports_Outgoing/");
            File[] arrFiles = dir.listFiles();
            List<File> lst = Arrays.asList(arrFiles);
            SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

            // пытаемся загрузить историю за X дней. готовим переменную dateBefore7Days
            Date d = new Date();
            Calendar cal = Calendar.getInstance();
            cal.setTime(d);
            writer.append(cal.getTime().toString());
            cal.add(Calendar.DATE, -365);
            Date dateBefore7Days = cal.getTime();

            for (int i = 1; i < lst.size(); i++) {
                File file = new File(lst.get(i).toString());

                final long lastModified = file.lastModified();
                //if (new Date(lastModified).after(dateBefore7Days)) {
                if (sdf.format(new Date(lastModified)).equals(sdf.format(d))) {
                    final String filename = file.getAbsolutePath();
                    Reader reader = null;

                    if (file.getName().toString().substring(0, 3).equals("EMF")) {
                        reader = new EMF_2_reader(filename);
                        reader.startread(con);
                    }
                    if (file.getName().toString().substring(0, 3).equals("ПВГ")) {
                        reader = new PVG_5_reader(filename);
                        reader.startread(con);
                    }
                    if (file.getName().toString().substring(0, 5).equals("ЧЗЛ_1")) {
                        reader = new CHZL1_4_reader(filename);
                        reader.startread(con);
                    }

                    if (file.getName().toString().substring(0, 3).equals("ПМО")) {
                        reader = new PMO_3_reader(filename);
                        reader.startread(con);
                    }
                    if (file.getName().toString().substring(0, 4).equals("ЧЗЛ7")) {
                        reader = new CHZL7_1_reader(filename);
                        reader.startread(con);
                    }

                    if (file.getName().toString().substring(0, 4).equals("МПНВ")) {
                        reader = new MPNV_6_reader(filename);
                        reader.startread(con);
                    }
                    if (file.getName().toString().substring(0, 6).equals("Policy")) {
                        reader = new PolicyTypes_8_reader(filename);
                        reader.startread(con);
                    }

                    if (file.getName().toString().substring(0, 3).equals("УЭК")) {
                        reader = new UEK_7_reader(filename);
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
}