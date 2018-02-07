package EISSOI;
import java.io.*;
import java.sql.Connection;
import java.util.Date;
import java.sql.PreparedStatement;
import java.text.SimpleDateFormat;
import java.util.Locale;

public class Deleter {
    public void deleter(Connection con, String tables, String file_date, String atServer) {
        try {
            // file_date приходит в виде строки yyyy/MM/dd
            SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd", Locale.US);
            // создаю дату и вставляю в нее текст
            Date docDate= format.parse(file_date);
            // т.к. в statement.setDate требуется передать sql.Date , то еще вот
            java.sql.Date endDate = new java.sql.Date(docDate.getTime());

            // exec нужен обязательно - т.к. запрос удаление данных с EISSOI отрабатывает удаленно( на 34 тачке !).
            String queryString = "exec( ' DELETE FROM ? WHERE file_date  =  ? ') ?";
            PreparedStatement statement = con.prepareStatement(queryString);
            statement.setString(1,tables);
            statement.setDate(2,endDate);
            statement.setString(3,atServer);
        } catch (Exception d) {
        }
    }
}
