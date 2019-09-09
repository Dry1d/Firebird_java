/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package otchet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author Dry1d
 */
public class ExcelWorker {

    //Текущая дата
    static LocalDate date = LocalDate.now();
    //Время начала опоздания
    static LocalTime time = LocalTime.of(8, 29, 30);
    static LocalDateTime datetime = LocalDateTime.of(date, time);
    //Время конца опоздания
    static LocalTime end_time = LocalTime.of(9, 10, 00);
    static LocalDateTime end_datetime = LocalDateTime.of(date, end_time);
    
     public static void setDate(LocalDate date) {
         ExcelWorker.date = date;
         //Переопределяем время начала и конца опозданий
         datetime = LocalDateTime.of(date, time);
         end_datetime = LocalDateTime.of(date, end_time);
     }

    // заполнение строки (rowNum) определенного листа (sheet)
    // данными  из dataModel созданного в памяти Excel файла
    private static void createSheetHeader(HSSFSheet sheet, int rowNum, DataModel dataModel) {

        Row row = sheet.createRow(rowNum);

        row.createCell(0).setCellValue(dataModel.getDate());
        row.createCell(1).setCellValue(dataModel.getTime());
        row.createCell(2).setCellValue(dataModel.getSt());
        row.createCell(3).setCellValue(dataModel.getDirection());
        row.createCell(4).setCellValue(dataModel.getFio());
    }

    private static void createSheetHeaderOp(HSSFSheet sheet, int rowNum, DataModel dataModel) {

        Row row = sheet.createRow(rowNum);

        row.createCell(0).setCellValue(dataModel.getDate());
        row.createCell(1).setCellValue(dataModel.getTime());
        row.createCell(2).setCellValue(dataModel.getSt());
        row.createCell(3).setCellValue(dataModel.getDirection());
        row.createCell(4).setCellValue(dataModel.getFio());
        row.createCell(5).setCellValue(dataModel.getPodr());
    }

    private static void createSheetHeaderNeyav(HSSFSheet sheet, int rowNum, DataModel dataModel) {

        Row row = sheet.createRow(rowNum);

        row.createCell(0).setCellValue(dataModel.getDate());
        row.createCell(1).setCellValue(dataModel.getFio());
        row.createCell(2).setCellValue(dataModel.getPodr());
    }

    public static void worker(String absolutefilepath, List<DataModel> dataModels) {

        // создание самого excel файла в памяти
        HSSFWorkbook workbook = new HSSFWorkbook();
//        // создание листа с названием "Просто лист"
//        HSSFSheet sheet = workbook.createSheet("Просто лист");

        // заполняем список какими-то данными
        List<DataModel> dataList = dataModels;

        for (DataModel dataModel : dataList) {
            HSSFSheet sheet;
            try {
                sheet = workbook.createSheet(dataModel.getPodr());
            } catch (Exception e) {
//                System.out.println("Лист "+dataModel.getPodr()+ " был создан ранее");
                sheet = workbook.getSheet(dataModel.getPodr());
            }

            // счетчик для строк
            int rowNum = 0;

            Row row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue("Дата");
            row.createCell(1).setCellValue("Время");
            row.createCell(2).setCellValue("Стойка");
            row.createCell(3).setCellValue("Направление");
            row.createCell(4).setCellValue("ФИО");

            List<DataModel> sheetDataModel = new ArrayList<>();
            for (DataModel dataModel0 : dataList) {
                if (dataModel.getPodr() == null ? dataModel0.getPodr() == null : dataModel.getPodr().equals(dataModel0.getPodr())) {
                    sheetDataModel.add(dataModel0);
                }

            }
            for (DataModel data : sheetDataModel) {
//                System.out.println(dataModel.getPodr());
//                System.out.println(data.getDate() + "|" + data.getTime() + "|" + data.getSt() + "|" + data.getDirection() + "|" + data.getFio());
                createSheetHeader(sheet, ++rowNum, data);
            }

        }
        //Создаём лист с опоздавшими детьми
        for (DataModel dataModel : dataList) {
            HSSFSheet sheet;
            try {
                sheet = workbook.createSheet("Опоздавшие");
            } catch (Exception e) {
//                System.out.println("Лист "+dataModel.getPodr()+ " был создан ранее");
                sheet = workbook.getSheet("Опоздавшие");
            }
//            HSSFSheet sheet = workbook.createSheet("Опоздавшие");

            // счетчик для строк
            int rowNum = 0;

            Row row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue("Дата");
            row.createCell(1).setCellValue("Время");
            row.createCell(2).setCellValue("Стойка");
            row.createCell(3).setCellValue("Направление");
            row.createCell(4).setCellValue("ФИО");
            row.createCell(5).setCellValue("Подразделение");

            List<DataModel> sheetDataModel0 = new ArrayList<>();
            for (DataModel dataModel1 : dataList) {

                try {
                    LocalDate date_event = LocalDate.parse(dataModel1.getDate());
                    LocalTime time_event = LocalTime.parse(dataModel1.getTime());
                    LocalDateTime event = LocalDateTime.of(date_event, time_event);

                    if (event.isAfter(datetime)) {
                        if (end_datetime.isAfter(event)) {
                            //Проверка условий вход-выход
                            //требуется добавить входил ли ученик/сотрудник сегодня ранее
                            if ("вход".equals(dataModel1.getDirection())) {
                                sheetDataModel0.add(dataModel1);
                            }
                        }
                    }
                } catch (Exception e) {

                }

            }
            for (DataModel data : sheetDataModel0) {
//                System.out.println("Опоздуны");
//                System.out.println(data.getDate() + "|" + data.getTime() + "|" + data.getSt() + "|" + data.getDirection() + "|" + data.getFio());
                createSheetHeaderOp(sheet, ++rowNum, data);
            }

        }

        //Создаём лист с неявившимися детьми
        for (DataModel dataModel : dataList) {
            HSSFSheet sheet;
            try {
                sheet = workbook.createSheet("Неявка");
            } catch (Exception e) {
//                System.out.println("Лист "+dataModel.getPodr()+ " был создан ранее");
                sheet = workbook.getSheet("Неявка");
            }

            // счетчик для строк
            int rowNum = 0;

            Row row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue("Дата");
            row.createCell(1).setCellValue("ФИО");
            row.createCell(2).setCellValue("Подразделение");

            List<DataModel> sheetDataModel1 = new ArrayList<>();
            for (DataModel dataModel2 : dataList) {

                if (dataModel2.getId() == 0 && !"Школа".equals(dataModel2.getPodr()) && !"тех.персонал".equals(dataModel2.getPodr()) && !"Педагоги".equals(dataModel2.getPodr()) && !"Мед_работники".equals(dataModel2.getPodr())) {
                    sheetDataModel1.add(dataModel2);
                }
            }
            for (DataModel data : sheetDataModel1) {
//                System.out.println("Опоздуны");
//                System.out.println(data.getDate() + "|" + data.getTime() + "|" + data.getSt() + "|" + data.getDirection() + "|" + data.getFio());
                createSheetHeaderNeyav(sheet, ++rowNum, data);
            }

        }

        // записываем созданный в памяти Excel документ в файл
        try (FileOutputStream out = new FileOutputStream(new File(absolutefilepath))) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel файл успешно создан!");
    }

}
