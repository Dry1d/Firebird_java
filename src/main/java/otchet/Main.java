/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package otchet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import javax.mail.*;

/**
 *
 * @author Dry1d
 */
public class Main {

    private static otchet.Sender sslSender;
    //инициализируем специальный объект Properties
    //типа Hashtable для удобной работы с данными
    public static Properties prop = new Properties();

    public static void main(String[] args) throws UnsupportedEncodingException, IOException, MessagingException {

        //Изменяемые параметры
        LocalDate date = LocalDate.now();
//        date = date.minusDays(5);

        String fileSeparator = System.getProperty("file.separator");
        String folder = "D:\\SCUD\\Otchet";
        String user = "SCD17_USER";
        String password = "scd17_password";
        String addr = "localhost";
        //Почта
        String subject = "";
        String content = "";
        String smtpHost = "";
        String mail_to = "";
        String mail_from = "";
        String mail_login = "";
        String mail_password = "";
        String smtpPort = "";
        boolean rew_flag = false;

        //Проверка условия наличия файла с параметрами, или аргументов
        
        if (new File(folder+fileSeparator+"properties.cfg").exists()) {
            FileInputStream fileInputStream;

            try {
                //обращаемся к файлу и получаем данные
                fileInputStream = new FileInputStream(folder+fileSeparator+"properties.cfg");
                prop.load(fileInputStream);

                folder = prop.getProperty("folder");
                user = prop.getProperty("user");
                password = prop.getProperty("password");
                addr = prop.getProperty("addr");
                //Почта
                smtpHost = prop.getProperty("smtpHost");
                mail_to = prop.getProperty("mail_to");
                mail_from = prop.getProperty("mail_from");
                mail_login = prop.getProperty("mail_login");
                mail_password = prop.getProperty("mail_password");
                smtpPort = prop.getProperty("smtpPort");

            } catch (IOException e) {
                System.out.println("Ошибка в программе: файл properties.cfg не обнаружен");
                e.printStackTrace();
            }
        } else {
            System.out.println("Properties не обнаружен!");
        }
        if (args.length != 0) {
            for (int i = 0; i < args.length; i++) {
                if (args[i].equals("-d")) {
                    String[] d = args[i + 1].split("-");
                    date = LocalDate.of(Integer.parseInt(d[0]), Integer.parseInt(d[1]), Integer.parseInt(d[2]));
                    System.out.println("Используем новую дату: " + date);
                } else if (args[i].equals("-f")) {
                    folder = args[i + 1];
                } else if (args[i].equals("-dbu")) {
                    user = args[i + 1];
                } else if (args[i].equals("-dbp")) {
                    password = args[i + 1];
                } else if (args[i].equals("-addr")) {
                    addr = args[i + 1];
                } else if (args[i].equals("-ml")) {
                    mail_login = args[i + 1];
                } else if (args[i].equals("-mp")) {
                    mail_password = args[i + 1];
                } else if (args[i].equals("-mt")) {
                    mail_to = args[i + 1];
                } else if (args[i].equals("-mf")) {
                    mail_from = args[i + 1];
                } else if (args[i].equals("-sH")) {
                    smtpHost = args[i + 1];
                } else if (args[i].equals("-sP")) {
                    smtpPort = args[i + 1];
                } else if (args[i].equals("-rw")) {
                    rew_flag = true;
                }
            }
        } else if(args.length == 0 && !new File(folder+fileSeparator+"properties.cfg").exists()){
            usage();
        }

        if (rew_flag) {
            System.out.println("Перезаписываем файл properties");
            prop.setProperty("folder", folder);
            prop.setProperty("user", user);
            prop.setProperty("password", password);
            prop.setProperty("addr", addr);
            prop.setProperty("smtpHost", smtpHost);
            prop.setProperty("mail_to", mail_to);
            prop.setProperty("mail_from", mail_from);
            prop.setProperty("mail_login", mail_login);
            prop.setProperty("mail_password", mail_password);
            prop.setProperty("smtpPort", smtpPort);
            
            
            prop.store(new FileWriter("."+fileSeparator+"properties.cfg"), "Перезаписываем файл properties");
        }
        System.out.println("Используется папка: " + folder + "\nАдрес: " + addr + "\nОтправляем с почты: " + mail_to + "\nНа почту: " + mail_from + "\nЧерез хост: " + smtpHost + ":" + smtpPort + "\nТема письма: " + subject + "\nТекст письма: " + content);

        //Заглушка для почты
        subject = "Отчет СКУД";
        content = date.toString();

        sslSender = new otchet.Sender(mail_to, mail_password, smtpHost, smtpPort);

        String databaseURL = "jdbc:firebirdsql://" + addr + ":3050/C:/ProgramData/PERCo-S-20/SCD17K.FDB?encoding=UTF8";
        String driverName = "org.firebirdsql.jdbc.FBDriver";

        String otchet_date = folder + fileSeparator + date.getYear() + "-" + date.getMonthValue() + "-" + date.getDayOfMonth() + ".xls";

//        String strSQL="select tabel_intermediadate.id_tb_in, datediff (second, cast('01.01.0001 00:00:00' as timestamp), tabel_intermediadate.date_pass ) seconds, tabel_intermediadate.date_pass, tabel_intermediadate.date_pass + tabel_intermediadate.time_pass timestamp_pass, tabel_intermediadate.type_pass, trim(staff.tabel_id) tabel_id, staff.last_name || ' ' || staff.first_name || ' ' || staff.middle_name fio, staff.id_staff from staff right outer join tabel_intermediadate on (staff.id_staff = tabel_intermediadate.staff_id)";
//        String strSQL="select tabel_intermediadate.date_pass + tabel_intermediadate.time_pass timestamp_pass, tabel_intermediadate.type_pass, configs_tree.display_name, staff.last_name || ' ' || staff.first_name || ' ' || staff.middle_name fio,  subdiv_ref.display_name from staff right outer join tabel_intermediadate on (staff.id_staff = tabel_intermediadate.staff_id) join staff_ref on (staff.id_staff = staff_ref.staff_id) join subdiv_ref on (staff_ref.subdiv_id = subdiv_ref.id_ref) join configs_tree on (tabel_intermediadate.config_tree_id = configs_tree.id_configs_tree) WHERE tabel_intermediadate.date_pass='2019-09-02' AND  subdiv_ref.id_ref = '5079'";
        //Запрос событий
        String strSQL = "select tabel_intermediadate.id_tb_in, tabel_intermediadate.date_pass timestamp_pass, tabel_intermediadate.time_pass timestamp_pass, tabel_intermediadate.type_pass, tabel_intermediadate.config_tree_id, staff.last_name || ' ' || staff.first_name || ' ' || staff.middle_name fio,  subdiv_ref.display_name , staff.id_staff from staff right outer join tabel_intermediadate on (staff.id_staff = tabel_intermediadate.staff_id) join staff_ref on (staff.id_staff = staff_ref.staff_id) join subdiv_ref on (staff_ref.subdiv_id = subdiv_ref.id_ref) join configs_tree on (tabel_intermediadate.config_tree_id = configs_tree.id_configs_tree) WHERE tabel_intermediadate.date_pass='" + date.getYear() + "-" + date.getMonthValue() + "-" + date.getDayOfMonth() + "'";
        //Запрос общего списка детей и сотрудников
        String op_strSQL = "select staff.id_staff, staff.last_name || ' ' || staff.first_name || ' ' || staff.middle_name fio,  subdiv_ref.display_name from staff right outer join staff_ref on (staff.id_staff = staff_ref.staff_id) join subdiv_ref on (staff_ref.subdiv_id = subdiv_ref.id_ref) ";

        try {
            // Инициализируемя Firebird JDBC driver.
            // Эта строка действительна только для Firebird.
            // Для других СУБД она будет немного видоизменена.
            Class.forName("org.firebirdsql.jdbc.FBDriver").newInstance();
        } catch (IllegalAccessException | InstantiationException | ClassNotFoundException ex) {
        }

        Connection conn = null;

        try {
            //Создаём подключение к базе данных
            conn = DriverManager.getConnection(databaseURL, user, password);
            if (conn == null) {
                System.err.println("Could not connect to database");
            }

            // Создаём класс, с помощью которого будут выполняться 
            // SQL запросы.
            Statement stmt = conn.createStatement();

            //Выполняем SQL запрос.
            //Список подразделений
            ResultSet rs = stmt.executeQuery(strSQL);

            // Смотрим количество колонок в результате SQL запроса.
//            int nColumnsCount = rs.getMetaData().getColumnCount();
            // Выводим результат.
            List<DataModel> dataModels = new ArrayList<>();

            while (rs.next()) {

                String direction = "";
                if ("6078".equals(rs.getObject(5).toString()) || "9852".equals(rs.getObject(5).toString())) {
                    direction = "вход";
                } else if ("6212".equals(rs.getObject(5).toString()) || "9718".equals(rs.getObject(5).toString())) {
                    direction = "выход";
                }
                if (!dataModels.isEmpty()) {
                    if (dataModels.get(dataModels.size() - 1).getId() != Integer.parseInt(rs.getObject(1).toString())) {
                        dataModels.add(new DataModel(Integer.parseInt(rs.getObject(1).toString()), rs.getObject(2).toString(), rs.getObject(3).toString(), rs.getObject(4).toString(), direction, rs.getObject(6).toString(), rs.getObject(7).toString(), rs.getObject(8).toString()));
                    } else {
//                        System.out.println("Событие уже было");
                    }
                } else {
                    dataModels.add(new DataModel(Integer.parseInt(rs.getObject(1).toString()), rs.getObject(2).toString(), rs.getObject(3).toString(), rs.getObject(4).toString(), direction, rs.getObject(6).toString(), rs.getObject(7).toString(), rs.getObject(8).toString()));
                }
            }

            //Создаём список тех, кто не пришел
            ResultSet rsOP = stmt.executeQuery(op_strSQL);
            //Список всех учеников и сотрудников без исключения
            List<DataModel> dMs = new ArrayList<>();

            while (rsOP.next()) {
//                LocalDateTime time = date.atTime(17, 0);
                dMs.add(new DataModel(0, date.toString(), "0", "0", "не явка", rsOP.getObject(2).toString(), rsOP.getObject(3).toString(), rsOP.getObject(1).toString()));
//                dMs.add(new DataModel(0, date.toString(), "0", "0", "не явка", rsOP.getObject(6).toString(), rsOP.getObject(7).toString(), rsOP.getObject(8).toString()));
            }

            //Нельзя проводить одновременно итерацию (перебор) коллекции и изменение ее элементов.
            //поэтому используем следующий код
            Iterator<DataModel> dMIterator = dMs.iterator();                    //создаем итератор
            while (dMIterator.hasNext()) {                                        //до тех пор, пока в списке есть элементы
                DataModel dM = dMIterator.next();                               //получаем следующий элемент
                dataModels.stream().filter((dataM) -> (dM.getId_staff().equals(dataM.getId_staff()))).forEachOrdered((_item) -> {
                    //                        System.out.println(dM.getId_staff() +" "+ dataM.getId_staff());
                    try {
                        dMIterator.remove();
                    } catch (Exception e) {

                    }
                });//                    System.out.println(dM.getId_staff()+" "+dataM.getId_staff());
//                    System.out.println(dM.getId_staff() == dataM.getId_staff());
            }

//            System.out.println(dMs.size());
            //Вносим список неявившихся в коллекцию
            dMs.forEach((op) -> {
                dataModels.add(op);
            });

            //Создаём документ excel
            ExcelWorker ew = new ExcelWorker();
            ew.setDate(date);
            ew.worker(otchet_date, dataModels);
            // Освобождаем ресурсы.
            stmt.close();
            conn.close();

            if (mail_login != "" ) {
                System.out.println("Идет отправка письма...");
                if (mail_from.contains(";")) {
                    String[] mails = mail_from.split(";");
                    for (String mail : mails) {
                        try {
                            sslSender.send(subject, content, mail_to, mail, otchet_date);
                        } catch (UnsupportedEncodingException e) {
                            System.out.println(e);
                        }
                    }
                } else {
                    try {
                        //Отправка на мыло отчета
                        sslSender.send(subject, content, mail_to, mail_from, otchet_date);
                    } catch (UnsupportedEncodingException e) {
                        System.out.println(e);
                    }
                }
                System.out.println("Сообщение успешно отправлено!");
            } else {
                System.out.println("Не заданы сетевые настройки для отправки отчета на электронную почту!");
            }
        } catch (SQLException ex) {
        }
    }

    private static void usage() {
        System.out.println("Генерация отчета от СКУД PERCo и отправка его по почте.");
        System.out.println("Первый запуск рекомендуется производить с параметрами:");
        System.out.println("java -jar PATH_TO_JAR_FILE\\Firebird_Java-0.0.1.jar -f D:\\\\SCUD\\Otchet\\ -addr localhost -ml login -mp <password> -mt ОТ_КОГО@DOMAIN -mf КОМУ@DOMAIN -sH smtp.yandex.ru -sP 465 -rw");
        System.out.println("-f D:\\\\SCUD\\Otchet\\                             -Рабочая папка");
        System.out.println("-addr localhost                                     -Адрес для подключения к БД Firebird");
        System.out.println("-ml login                                           -Логин от почты");
        System.out.println("-mp <password>                                      -Пароль от почты");
        System.out.println("-mt ОТ_КОГО@DOMAIN                                  -От кого");
        System.out.println("-mf КОМУ@DOMAIN                                     -Кому, возможно перечисление, например: почта1;почта2;почта3");
        System.out.println("-sH smtp.yandex.ru "
                         + "-sP 465                                             -Параметры для подключения к серверу smtp");
        System.out.println("-rw                                                 -Перезапись файла properties.cfg параметрами из аргументов");
        System.out.println("                                                    -файл properties.cfg формируется в папке folder");
        System.out.println("По умолчанию производилась отправка с yandex почты по протоколу SSL, если не работает, потребуется создавать класс TLS_Sender или другой");
    }

//    private static String bscrypt_salt(String mail_password) {
//        String generatedSecuredPasswordHash = BCrypt.hashpw(mail_password, BCrypt.gensalt(12));
//        System.out.println(generatedSecuredPasswordHash);
//        return(generatedSecuredPasswordHash);
//    }
}
