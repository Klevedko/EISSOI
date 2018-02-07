package EISSOI;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.ConnectException;
import java.sql.Connection;
import java.util.Iterator;

import static EISSOI.App.*;

public class UEK_7_reader extends Reader {
    public UEK_7_reader(String fileName,String target) {
        super(fileName,target);
    }

    public void startread(Connection con) {
        try {
            InputStream excelFile = new BufferedInputStream(new FileInputStream(new File(filename)));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            writer.append("---------" + filename + "----------");

            // Удаление из ReportData
            deleteSql="exec ( 'delete from ReportAnalize_UEK_History_java where file_date = ''" + target + "''')";
            App.connecting(con, filename, deleteSql);

            // Удаление из EISSOI
            deleteSqlEissoi = deleteSql.replaceAll("ReportAnalize_UEK_History_java", "erz_exp.dbo.ReportAnalize_UEK_History_java");
            deleteSqlEissoi = deleteSqlEissoi + " at [MOS-EISSOI-03]";
            App.connecting(con, filename, deleteSqlEissoi);

            while (iterator.hasNext()) {
                sql = "exec(' insert into ReportAnalize_UEK_History_java ([date_insert]\n" +
                        "      ,[RF_part]\n" +
                        "      ,[Code]\n" +
                        "      ,[count_of_dead_on_date2]\n" +
                        "      ,[uek_total_on_date_1]\n" +
                        "      ,[uek_total_on_date_2]\n" +
                        "      ,[from_uek_on_ERZ_on_date_1]\n" +
                        "      ,[from_uek_on_ERZ_on_date_2]\n" +
                        "      ,[Title]\n" +
                        "      ,[date1]\n" +
                        "      ,[date2]\n" +
                        "      ,[file__name], [file_date] ) select " + rowInsert;
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    currentCell.setCellType(Cell.CELL_TYPE_STRING);
                    DataFormatter formatter = new DataFormatter();
                    if (currentCell.getColumnIndex() < 17) {
                        if (currentRow.getRowNum() >= 7) {
                            if (currentCell.getColumnIndex() > 4) {
                                sql = sql + (formatter.formatCellValue(currentCell).isEmpty() ? "0 " :
                                        formatter.formatCellValue(currentCell)) + ", ";
                            }
                            if (currentCell.getColumnIndex() == 4)
                                sql = sql + "''" + formatter.formatCellValue(currentCell) + "''" + ", ";
                        } else {
                            sql = "" + System.getProperty("line.separator");
                            if (currentCell.getAddress().toString().equals("C2"))
                                Title = "''" + formatter.formatCellValue(currentCell) + "''";
                            if (currentCell.getAddress().toString().equals("H6"))
                                date1 = "''" + formatter.formatCellValue(currentCell) + "''";
                            if (currentCell.getAddress().toString().equals("G6"))
                                date2 = "''" + formatter.formatCellValue(currentCell) + "''";
                        }
                    }
                }
                sql = sql + Title + ", " + date1 + ", " + date2 + ", " + "''" + filename + "'',''" + target + "''')";
                sqlEISSOI = sql.replaceAll("ReportAnalize_UEK_History_java", "erz_exp.dbo.ReportAnalize_UEK_History_java");
                sqlEISSOI = sqlEISSOI + " at [MOS-EISSOI-03]";

                App.connecting(con, filename, sql);
                App.connecting(con, filename, sqlEISSOI);
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
