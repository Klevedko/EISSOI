package EISSOI;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.ConnectException;
import java.sql.Connection;
import java.util.Iterator;

import static EISSOI.App.*;

public class PolicyTypes_8_reader extends Reader {
    public PolicyTypes_8_reader(String fileName,String target) {
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
            deleteSql="exec ( 'delete from ReportAnalize_PolicyTypes_History_java where file_date = ''" + target + "''')";
            App.connecting(con, filename, deleteSql);

            // Удаление из EISSOI
            deleteSqlEissoi = deleteSql.replaceAll("ReportAnalize_PolicyTypes_History_java", "erz_exp.dbo.ReportAnalize_PolicyTypes_History_java");
            deleteSqlEissoi = deleteSqlEissoi + " at [MOS-EISSOI-03]";
            App.connecting(con, filename, deleteSqlEissoi);

            while (iterator.hasNext()) {
                sql = "exec(' insert into ReportAnalize_PolicyTypes_History_java ([date_insert]\n" +
                        "      ,[RF_part]\n" +
                        "      ,[Code]\n" +
                        "      ,[count_total_in_SZ_ERZ]\n" +
                        "      ,[colunt_new_polices]\n" +
                        "      ,[count_unusal_documents]\n" +
                        "      ,[count_old_polices]\n" +
                        "      ,[vcount_UEK]\n" +
                        "      ,[Parent]\n" +
                        "      ,[file__name], [file_date] ) select " + rowInsert;
                System.out.println(sql);
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    currentCell.setCellType(Cell.CELL_TYPE_STRING);
                    DataFormatter formatter = new DataFormatter();
                    if (currentCell.getColumnIndex() < 17) {
                        if (currentCell.getAddress().toString().equals("D2"))
                            Title = "''" + formatter.formatCellValue(currentCell) + "''";
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
                sqlEISSOI = sql.replaceAll("ReportAnalize_PolicyTypes_History_java", "erz_exp.dbo.ReportAnalize_PolicyTypes_History_java");
                sqlEISSOI=sqlEISSOI+ " at [MOS-EISSOI-03]";

                App.connecting(con,filename,sql);
                App.connecting(con,filename,sqlEISSOI);
            }
        }  catch (FileNotFoundException e) {
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
