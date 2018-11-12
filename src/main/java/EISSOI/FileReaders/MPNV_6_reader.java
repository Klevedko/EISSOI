package EISSOI.FileReaders;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.ConnectException;
import java.sql.Connection;
import java.util.Iterator;

import static EISSOI.App.*;
import static EISSOI.Psql.sqlConn.connecting;

public class MPNV_6_reader extends EISSOI.AbstractReader.Reader {
    public MPNV_6_reader(String fileName, String target) {
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
            deleteSql = "exec ( 'delete from ReportAnalize_MPNV_History_java where file_date = ''" + target + "''')";
            connecting(con, filename, deleteSql);

            // Удаление из EISSOI
            deleteSqlEissoi = deleteSql.replaceAll("ReportAnalize_MPNV_History_java", "erz_exp.dbo.ReportAnalize_MPNV_History_java");
            deleteSqlEissoi = deleteSqlEissoi + " at [MOS-EISSOI-03]";
            connecting(con, filename, deleteSqlEissoi);

            while (iterator.hasNext()) {
                sql = "exec(' insert into ReportAnalize_MPNV_History_java ([date_insert]\n" +
                        "      ,[RF_part]\n" +
                        "      ,[Code]\n" +
                        "      ,[pervichnyh]\n" +
                        "      ,[fixed]\n" +
                        "      ,[total]\n" +
                        "      ,[not_found_ZL_500]\n" +
                        "      ,[not_found_MO_265]\n" +
                        "      ,[MO_other_territory_541]\n" +
                        "      ,[no_connect_MO_542]\n" +
                        "      ,[not_found_doctor_543]\n" +
                        "      ,[doctor_doesnt_work_544]\n" +
                        "      ,[date_conflict_547]\n" +
                        "      ,[post_conflict_545]\n" +
                        "      ,[More_than_2_connections_546]\n" +
                        "      ,[Attached_total]\n" +
                        "      ,[Attached_to_doctors]\n" +
                        "      ,[Attached_to_Junior_nurses]\n" +
                        "      ,[Attached_to_2_doctors]\n" +
                        "      ,[DTotal]\n" +
                        "      ,[DTotal_less_5_ZL]\n" +
                        "      ,[DTotalMax_more_3000_ZL]\n" +
                        "      ,[DAverage]\n" +
                        "      ,[DMinimum]\n" +
                        "      ,[DMaximum]\n" +
                        "      ,[Title]\n" +
                        "      ,[file__name], [file_date] ) select " + rowInsert;
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    currentCell.setCellType(Cell.CELL_TYPE_STRING);
                    DataFormatter formatter = new DataFormatter();
                    if (currentCell.getColumnIndex() < 27) {
                        if (currentCell.getAddress().toString().equals("B2")) {
                            Title = "''" + formatter.formatCellValue(currentCell) + "''";
                        }
                        if (currentRow.getRowNum() >= 6) {
                            if (currentCell.getColumnIndex() > 3) {
                                sql = sql + (formatter.formatCellValue(currentCell).isEmpty() ? "0 " :
                                        formatter.formatCellValue(currentCell)) + ", ";
                            }
                            if (currentCell.getColumnIndex() == 3)
                                sql = sql + "''" + formatter.formatCellValue(currentCell) + "''" + ", ";
                        }
                    }
                }
                sql = sql + Title + ", " + "''" + filename + "'',''" + target + "''')";
                sqlEISSOI = sql.replaceAll("ReportAnalize_MPNV_History_java", "erz_exp.dbo.ReportAnalize_MPNV_History_java");
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
