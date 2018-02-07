package EISSOI;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public abstract class Reader {
    public static String sql = "";
    public static String deleteSql = "";
    public static String deleteSqlEissoi = "";
    public String Title = "";
    public String date1 = "";
    public String date2 = "";
    public DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM/dd/yyyy");
    public LocalDateTime now = LocalDateTime.now();
    public String rowInsert = "''" + dtf.format(now) + "'', ";
    public String form8 = "";
    public String sZ_ERZ = "";
    public String Index_dead_in_period = "";
    public String Index_dead_in_period_ROSSTAT = "";
    public String Index_quartal = "";
    public String Index_Working= "";
    public final String filename;
    public final String target;
    public static String  sqlEISSOI="";
    //public Statement st = null;
    Reader(String filename, String target){
        this.filename=filename;
        this.target=target;
    }
    public abstract void startread(Connection con);
}
