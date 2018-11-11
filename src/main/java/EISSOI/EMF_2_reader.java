package EISSOI;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.ConnectException;
import java.sql.Connection;
import java.util.Iterator;

import static EISSOI.App.*;

public class EMF_2_reader extends Reader {

    public EMF_2_reader(String fileName, String target) {
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
            deleteSql="exec ( 'delete from ReportAnalize_EMF_history_java where file_date = ''" + target + "''')";
            App.connecting(con, filename, deleteSql);

            // Удаление из EISSOI
            deleteSqlEissoi = deleteSql.replaceAll("ReportAnalize_EMF_history_java", "erz_exp.dbo.ReportAnalize_EMF_history_java");
            deleteSqlEissoi = deleteSqlEissoi + " at [MOS-EISSOI-03]";
            App.connecting(con, filename, deleteSqlEissoi);

            while (iterator.hasNext()) {
                sql = "exec ('insert into ReportAnalize_EMF_history_java ( [date_insert]\n" +
                        "      ,[RF]\n" +
                        "      ,[code]\n" +
                        "      ,[in_ERZ_current_date]\n" +
                        "      ,[see_dday1]\n" +
                        "      ,[see_dday2]\n" +
                        "      ,[see_dday3]\n" +
                        "      ,[see_dday4]\n" +
                        "      ,[opened_insurance_and_dead]\n" +
                        "      ,[DPFS]\n" +
                        "      ,[Ended_plan_date]\n" +
                        "      ,[no_info_about_UDL]\n" +
                        "      ,[expired_VC]\n" +
                        "      ,[in_SZ_ERZ_with_UEK]\n" +
                        "      ,[FUO_total]\n" +
                        "      ,[FUO_done_and_expired]\n" +
                        "      ,[see_dday5]\n" +
                        "      ,[see_dday6_do_we_have_info_TFOMS_to_SZ_ERZ]\n" +
                        "      ,[see_dday6_QUARTAL_WORK]\n" +
                        "      ,[see_dday6_QUARTAL_working_in_other_subject_RF]\n" +
                        "      ,[see_dday6_QUARTAL_insured_in_other_TC]\n" +
                        "      ,[ZL_TOTAL]\n" +
                        "      ,[to_one_MO]\n" +
                        "      ,[to_one_MO_in_other_territory]\n" +
                        "      ,[TOTAL_multilink_MO]\n" +
                        "      ,[TOTAL_multilink_MO_in_other_territory]\n" +
                        "      ,[Attach_to_med_personal]\n" +
                        "      ,[CountDoctor]\n" +
                        "      ,[AttachAVG]\n" +
                        "      ,[AttachMIN]\n" +
                        "      ,[AttachMAX]\n" +
                        "      ,[wrong_sex_or_age]\n" +
                        "      ,[wrong_TFOMS_code]\n" +
                        "      ,[with_error_in_line]\n" +
                        "      ,[Title]\n" +
                        "      ,[dday1]\n" +
                        "      ,[dday2]\n" +
                        "      ,[dday3]\n" +
                        "      ,[dday4]\n" +
                        "      ,[dday5]\n" +
                        "      ,[dday6]\n" +
                        "      ,[file__name], [file_date] ) select " + rowInsert;
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    currentCell.setCellType(Cell.CELL_TYPE_STRING);
                    DataFormatter formatter = new DataFormatter();
                    if (currentCell.getAddress().toString().equals("D2")) {
                        Title = "''" + formatter.formatCellValue(currentCell) + "''";
                    }
                    if (currentCell.getAddress().toString().equals("G4")) {
                        form8 = "''" + formatter.formatCellValue(currentCell) + "''";
                    }
                    if (currentCell.getAddress().toString().equals("H4")) {
                        sZ_ERZ = "''" + formatter.formatCellValue(currentCell) + "''";
                    }
                    if (currentCell.getAddress().toString().equals("J4")) {
                        Index_dead_in_period = "''" + formatter.formatCellValue(currentCell) + "''";
                    }
                    if (currentCell.getAddress().toString().equals("K4")) {
                        Index_dead_in_period_ROSSTAT = "''" + formatter.formatCellValue(currentCell) + "''";
                    }
                    if (currentCell.getAddress().toString().equals("T4")) {
                        Index_quartal = "''" + formatter.formatCellValue(currentCell) + "''";
                    }
                    if (currentCell.getAddress().toString().equals("U4")) {
                        Index_Working = "''" + formatter.formatCellValue(currentCell) + "''";
                    }
                    if (!currentCell.getAddress().toString().substring(0, 1).equals("I") && currentRow.getRowNum() < 104) {
                        if (currentRow.getRowNum() >= 8) {
                            if (currentCell.getColumnIndex() == 3 || currentCell.getAddress().toString().substring(0, 1).equals("U"))
                                sql = sql + "''" + formatter.formatCellValue(currentCell) + "''" + ", ";
                            else
                                sql = sql + (formatter.formatCellValue(currentCell).isEmpty() ? "0" : formatter.formatCellValue(currentCell)) + ", ";
                        }
                    }

                }

                sql = sql + Title + "," + form8 + "," + sZ_ERZ + "," + Index_dead_in_period + ","
                        + Index_dead_in_period_ROSSTAT + "," + Index_quartal + "," + Index_Working + ",''" + filename + "'',''" + target + "''')";
                sqlEISSOI = sql.replaceAll("ReportAnalize_EMF_history_java", "erz_exp.dbo.ReportAnalize_EMF_history_java");
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
