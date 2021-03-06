package EISSOI.FileReaders;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.ConnectException;
import java.sql.Connection;
import java.util.Iterator;

import static EISSOI.App.writer;
import static EISSOI.Psql.sqlConn.connecting;

public class OVP_reader extends EISSOI.AbstractReader.Reader {
    public OVP_reader(String fileName, String target) {
        super(fileName, target);
    }

    public void startread(Connection con) {
        try {
            InputStream excelFile = new BufferedInputStream(new FileInputStream(new File(filename)));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            writer.append("---------" + filename + "----------");

            // Удаление из ReportData
            deleteSql = "exec ( 'delete from ReportAnalize_OVP_History_java where file_date = ''" + target + "''')";
            connecting(con, filename, deleteSql);

            // Удаление из EISSOI
            deleteSqlEissoi = deleteSql.replaceAll("ReportAnalize_OVP_History_java", "erz_exp.dbo.ReportAnalize_OVP_History_java");
            deleteSqlEissoi = deleteSqlEissoi + " at [MOS-EISSOI-03]";
            connecting(con, filename, deleteSqlEissoi);

            while (iterator.hasNext()) {
                sql = "exec(' insert into ReportAnalize_OVP_History_java ([date_insert]\n" +
                        "      ,[RF_part]\n" +
                        "      ,[Code]\n" +
                        "      ,[Request_Given_B]\n" +
                        "      ,[Request_Given_E]\n" +
                        "      ,[Request_Authorized_B]\n" +
                        "      ,[Request_Authorized_E]\n" +
                        "      ,[Request_Taken_to_execution_B]\n" +
                        "      ,[Request_Taken_to_execution_E]\n" +
                        "      ,[Request_Denied_execution_B]\n" +
                        "      ,[Request_Denied_execution_E]\n" +
                        "      ,[polises_Made_B]\n" +
                        "      ,[polises_Made_E]\n" +
                        "      ,[polises_Received_in_TFOMS_B]\n" +
                        "      ,[polises_Received_in_TFOMS_E]\n" +
                        "      ,[Title]\n" +
                        "      ,[file__name], [file_date] )  select " + rowInsert;
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getColumnIndex() <= 16) {
                        currentCell.setCellType(Cell.CELL_TYPE_STRING);
                        DataFormatter formatter = new DataFormatter();
                        if (currentCell.getAddress().toString().equals("B2")) {
                            Title = "''" + formatter.formatCellValue(currentCell) + "''";
                        }

                        if (currentRow.getRowNum() >= 7) {
                            if (currentCell.getColumnIndex() > 3) {
                                sql = sql + (formatter.formatCellValue(currentCell).isEmpty() ? "0 " :
                                        formatter.formatCellValue(currentCell)) + ", ";
                            }
                            if (currentCell.getColumnIndex() == 3)
                                sql = sql + "''" + formatter.formatCellValue(currentCell) + "''" + ", ";
                        }
                    }
                }
                sql = sql + Title + "," + "''" + filename + "'',''" + target + "''')";
                sqlEISSOI = sql.replaceAll("ReportAnalize_OVP_History_java", "erz_exp.dbo.ReportAnalize_OVP_History_java");
                sqlEISSOI = sqlEISSOI + " at [MOS-EISSOI-03]";

                connecting(con, filename, sql);
                connecting(con, filename, sqlEISSOI);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (ConnectException c) {
            try {
                writer.append('\n');
                writer.append(c.getMessage().toString());
            } catch (Exception ee) {
            }
        } catch (Exception e) {
            try {
                writer.append('\n');
                writer.append(e.getMessage().toString());
                writer.flush();
                e.printStackTrace();
            } catch (Exception ee) {
            }
        }
    }
}
