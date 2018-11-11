package EISSOI.Psql;

import java.sql.*;
import java.lang.String;

import static EISSOI.App.writer;

public class sqlConn {

    Statement st = null;

    public void connecting(Connection con, String f, String QUERRRY) {
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