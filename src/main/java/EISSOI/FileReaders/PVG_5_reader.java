package EISSOI.FileReaders;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.ConnectException;
import java.sql.*;
import java.util.Iterator;

import static EISSOI.App.*;
import static EISSOI.Psql.sqlConn.connecting;

public class PVG_5_reader extends EISSOI.AbstractReader.Reader {
    public PVG_5_reader(String fileName, String target) {
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
            deleteSql = "exec ( 'delete from ReportAnalize_AttachmentCountDoctor_history_java where file_date = ''" + target + "''')";
            connecting(con, filename, deleteSql);

            // Удаление из EISSOI
            deleteSqlEissoi = deleteSql.replaceAll("ReportAnalize_AttachmentCountDoctor_history_java", "erz_exp.dbo.ReportAnalize_AttachmentCountDoctor_history_java");
            deleteSqlEissoi = deleteSqlEissoi + " at [MOS-EISSOI-03]";
            connecting(con, filename, deleteSqlEissoi);

            while (iterator.hasNext()) {
                sql = "exec(' insert into ReportAnalize_AttachmentCountDoctor_history_java ([date_insert]\n" +
                        "      ,[RF_part]\n" +
                        "      ,[Code]\n" +
                        "      ,[from200_to900]\n" +
                        "      ,[less200]\n" +
                        "      ,[less_100]\n" +
                        "      ,[less50]\n" +
                        "      ,[less5]\n" +
                        "      ,[Title]\n" +
                        "      ,[file__name], [file_date] )  select " + rowInsert;
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getColumnIndex() <= 9) {
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
                sqlEISSOI = sql.replaceAll("ReportAnalize_AttachmentCountDoctor_history_java", "erz_exp.dbo.ReportAnalize_AttachmentCountDoctor_history_java");
                sqlEISSOI = sqlEISSOI + " at [MOS-EISSOI-03]";
                // передаем соединение и вставляем строку

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
