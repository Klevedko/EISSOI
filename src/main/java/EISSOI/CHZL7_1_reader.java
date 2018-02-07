package EISSOI;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.ConnectException;
import java.sql.Connection;
import java.util.Iterator;
import static EISSOI.App.*;
public class CHZL7_1_reader extends Reader{
    public CHZL7_1_reader(String fileName,String target) {
        super(fileName,target);
    }
    public void startread(Connection con) {
        try {
            InputStream excelFile = new BufferedInputStream(new FileInputStream(new File(filename)));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            writer.append("---------" + filename + "----------");

            while (iterator.hasNext()) {
                sql = "exec(' insert into ReportEmploymentMonitoringPod_history_java ([date_insert]\n" +
                        "      ,[СУБЪЕКТ РФ]\n" +
                        "      ,[Total]\n" +
                        "      ,[РОССИЙСКАЯ ФЕДЕРАЦИЯ]\n" +
                        "      ,[ЦЕНТРАЛЬНЫЙ ФЕДЕРАЛЬНЫЙ ОКРУГ]\n" +
                        "      ,[Белгородская область - 14000]\n" +
                        "      ,[Брянская область - 15000]\n" +
                        "      ,[Владимирская область - 17000]\n" +
                        "      ,[Воронежская область - 20000]\n" +
                        "      ,[Ивановская область - 24000]\n" +
                        "      ,[Калужская область - 29000]\n" +
                        "      ,[Костромская область - 34000]\n" +
                        "      ,[Курская область - 38000]\n" +
                        "      ,[Липецкая область - 42000]\n" +
                        "      ,[Московская область - 46000]\n" +
                        "      ,[Орловская область - 54000]\n" +
                        "      ,[Рязанская область - 61000]\n" +
                        "      ,[Смоленская область - 66000]\n" +
                        "      ,[Тамбовская область - 68000]\n" +
                        "      ,[Тверская область - 28000]\n" +
                        "      ,[Тульская область - 70000]\n" +
                        "      ,[Ярославская область - 78000]\n" +
                        "      ,[г.Москва - 45000]\n" +
                        "      ,[СЕВЕРО-ЗАПАДНЫЙ ФЕДЕРАЛЬНЫЙ ОКРУГ]\n" +
                        "      ,[Республика Карелия - 86000]\n" +
                        "      ,[Республика Коми - 87000]\n" +
                        "      ,[Архангельская область - 11000]\n" +
                        "      ,[Вологодская область - 19000]\n" +
                        "      ,[Калининградская область - 27000]\n" +
                        "      ,[Ленинградская область - 41000]\n" +
                        "      ,[Мурманская область - 47000]\n" +
                        "      ,[Новгородская область - 49000]\n" +
                        "      ,[Псковская область - 58000]\n" +
                        "      ,[г.Санкт-Петербург - 40000]\n" +
                        "      ,[Ненецкий АО - 11100]\n" +
                        "      ,[ЮЖНЫЙ ФЕДЕРАЛЬНЫЙ ОКРУГ]\n" +
                        "      ,[Республика Адыгея - 79000]\n" +
                        "      ,[Республика Калмыкия - 85000]\n" +
                        "      ,[Краснодарский край - 3000]\n" +
                        "      ,[Астраханская область - 12000]\n" +
                        "      ,[Волгоградская область - 18000]\n" +
                        "      ,[Ростовская область - 60000]\n" +
                        "      ,[г. Севастополь - 67000]\n" +
                        "      ,[Республика Крым - 35000]\n" +
                        "      ,[СЕВЕРО-КАВКАЗСКИЙ ФЕДЕРАЛЬНЫЙ ОКРУГ]\n" +
                        "      ,[Кабардино-Балкарская Республика - 83000]\n" +
                        "      ,[Карачаево-Черкесская Республика - 91000]\n" +
                        "      ,[Республика Дагестан - 82000]\n" +
                        "      ,[Республика Ингушетия - 26000]\n" +
                        "      ,[Республика Северная Осетия-Алания - 90000]\n" +
                        "      ,[Чеченская Республика - 96000]\n" +
                        "      ,[Ставропольский край - 7000]\n" +
                        "      ,[ПРИВОЛЖСКИЙ ФЕДЕРАЛЬНЫЙ ОКРУГ]\n" +
                        "      ,[Республика Башкортостан - 80000]\n" +
                        "      ,[Республика Марий Эл - 88000]\n" +
                        "      ,[Республика Мордовия - 89000]\n" +
                        "      ,[Республика Татарстан - 92000]\n" +
                        "      ,[Удмуртская Республика - 94000]\n" +
                        "      ,[Чувашская Республика - 97000]\n" +
                        "      ,[Пермский край - 57000]\n" +
                        "      ,[Кировская область - 33000]\n" +
                        "      ,[Нижегородская область - 22000]\n" +
                        "      ,[Оренбургская область - 53000]\n" +
                        "      ,[Пензенская область - 56000]\n" +
                        "      ,[Самарская область - 36000]\n" +
                        "      ,[Саратовская область - 63000]\n" +
                        "      ,[Ульяновская область - 73000]\n" +
                        "      ,[УРАЛЬСКИЙ ФЕДЕРАЛЬНЫЙ ОКРУГ]\n" +
                        "      ,[Курганская область - 37000]\n" +
                        "      ,[Свердловская область - 65000]\n" +
                        "      ,[Тюменская область - 71000]\n" +
                        "      ,[Челябинская область - 75000]\n" +
                        "      ,[Ханты-Мансийский АО - 71100]\n" +
                        "      ,[Ямало-Ненецкий АО - 71140]\n" +
                        "      ,[СИБИРСКИЙ ФЕДЕРАЛЬНЫЙ ОКРУГ]\n" +
                        "      ,[Республика Алтай - 84000]\n" +
                        "      ,[Республика Бурятия - 81000]\n" +
                        "      ,[Республика Тыва - 93000]\n" +
                        "      ,[Республика Хакасия - 95000]\n" +
                        "      ,[Алтайский край - 1000]\n" +
                        "      ,[Забайкальский край - 76000]\n" +
                        "      ,[Красноярский край - 4000]\n" +
                        "      ,[Иркутская область - 25000]\n" +
                        "      ,[Кемеровская область - 32000]\n" +
                        "      ,[Новосибирская область - 50000]\n" +
                        "      ,[Омская область - 52000]\n" +
                        "      ,[Томская область - 69000]\n" +
                        "      ,[ДАЛЬНЕВОСТОЧНЫЙ ФЕДЕРАЛЬНЫЙ ОКРУГ]\n" +
                        "      ,[Республика Саха (Якутия) - 98000]\n" +
                        "      ,[Камчатский край - 30000]\n" +
                        "      ,[Приморский край - 5000]\n" +
                        "      ,[Хабаровский край - 8000]\n" +
                        "      ,[Амурская область - 10000]\n" +
                        "      ,[Магаданская область - 44000]\n" +
                        "      ,[Сахалинская область - 64000]\n" +
                        "      ,[Еврейская АО - 99000]\n" +
                        "      ,[Чукотский АО - 77000]\n" +
                        "      ,[БАЙКОНУР]\n" +
                        "      ,[г.Байконур - 55000]\n" +
                        "      ,[Title]\n" +
                        "      ,[file__name], [file_date] ) select " + rowInsert;
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    currentCell.setCellType(Cell.CELL_TYPE_STRING);
                    DataFormatter formatter = new DataFormatter();
                    if (currentCell.getAddress().toString().equals("B3")) {
                        Title = "''" + formatter.formatCellValue(currentCell) + "''";
                    }
                    if (!currentCell.getAddress().toString().substring(0, 2).equals("BS")
                            && !currentCell.getAddress().toString().substring(0, 2).equals("BT")
                            && !currentCell.getAddress().toString().substring(0, 2).equals("CY")
                            && !currentCell.getAddress().toString().substring(0, 2).equals("CX")) {
                        if (currentRow.getRowNum() >= 6) {
                            if (currentCell.getColumnIndex() >= 2) {
                                sql = sql + (formatter.formatCellValue(currentCell).isEmpty() ? "0" :
                                        formatter.formatCellValue(currentCell)) + ", ";
                            } else {
                                sql = sql + "''" + formatter.formatCellValue(currentCell) + "''" + ", ";
                            }
                        }
                    }
                }
                sql=sql+Title+"," + "''" + filename + "'',''" + target + "''')";
                sqlEISSOI = sql.replaceAll("ReportEmploymentMonitoringPod_history_java", "erz_exp.dbo.ReportEmploymentMonitoringPod_history_java");
                sqlEISSOI=sqlEISSOI+ " at [MOS-EISSOI-03]";

                App.connecting(con,filename,sql);
                App.connecting(con,filename,sqlEISSOI);
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

