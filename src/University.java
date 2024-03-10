
import java.awt.*;
import java.awt.event.InputEvent;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.AbstractMap.SimpleEntry;


public class University {
    public static void main(String[] args) {

        List<Student> studentList;
        List<Teacher> teacherList;
        boolean exit = false;
        Scanner scanner = new Scanner(System.in);
            cls();
            System.out.println("Вас вітає програма обліку відвідування та успішності студентів!\n");
            while (!exit) {
                System.out.println("\n ГОЛОВНЕ МЕНЮ:");
                System.out.println("---------------------------------------------------");
                System.out.println("1 - Коригувати відомості про студента");
                System.out.println("2 - Коригувати відомості про викладачів");
                System.out.println("3 - Перейти до журналу");
                System.out.println("4 - Вийти з програми");
                System.out.println("---------------------------------------------------");
                System.out.print("Введіть ваш вибір: ");
                while (!scanner.hasNextInt()) {
                    System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                    scanner.next(); // Переміщуємо курсор вводу вперед
                }
                // Якщо користувач ввів ціле число, зчитуємо його
                int Gen_num = scanner.nextInt();
                scanner.nextLine();
                    switch (Gen_num) {
                    case (1):
                        cls();
                        studentList = Student.readStudent();
                        Student.printStudent(studentList);
                        submenuStudent(scanner);
                       break;
                    case (2):
                        cls();
                        teacherList = Teacher.readTeacher();
                        Teacher.printTeacher(teacherList);
                        submenuTeacher(scanner);
                        break;
                    case (3):
                        cls();
                        // Створюємо сутність класу Група
                        Group newgroup = new Group();
                        // Створюємо сутність класу Предмет
                        Subject newsubject = new Subject (scanner, newgroup.groupNumber);
                        // Створюємо назву листа таблиці EXEL
                        String codAbbreviation = Teacher.generateShortSubject(newsubject.getSubject());
                        String cod = codAbbreviation + " " + newgroup.groupNumber;
                        // перевіряємо чи вже є лист з назвою  cod
                       try (FileInputStream fis = new FileInputStream("resources/Univer.xlsx");
                             Workbook workbook = WorkbookFactory.create(fis)) {
                           boolean lessOld_or_New = false;
                           boolean sheetExists = false;
                           for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                               if (workbook.getSheetName(i).equalsIgnoreCase(cod)) {
                                   sheetExists = true;
                                   break;
                               }
                           }
                           if (sheetExists) {// коли лист EXEL існує
                               //Створюємо сутність класу Заняття
                               List<SimpleEntry<String, String>> oldDataList = ExelMagazinDialog.readData (cod);
                               Lesson newlesson = new Lesson (newsubject.subject, oldDataList);
                               SimpleEntry simpl = new SimpleEntry<>(newlesson.getDate(), newlesson.getType());
                               //перевіряємо чи це заняття нове чи вже існуюче
                               if (oldDataList.contains(simpl)) {
                                   //якщо вже проводилось заняття
                                   lessOld_or_New = true;
                                   List<Attendance> oldAttendanceList = ExelMagazinDialog.readlistGradesMagazin(cod, newlesson);
                                   submenuMagazine(scanner, cod, newlesson, newgroup, oldAttendanceList, lessOld_or_New);

                               } else {
                                   lessOld_or_New = false;
                                   //зчитуємо сутність класу Група з EXEL
                                   Group oldGroup = ExelMagazinDialog.readGroup(cod);
                                   //перевіряємо чи перелік студентів не змінився
                                   if (oldGroup.fullNameStudents.size() == newgroup.fullNameStudents.size()) {
                                       if (oldGroup.fullNameStudents.containsAll(newgroup.fullNameStudents)) {

                                           Lesson tempLesson = new Lesson(oldDataList.get(oldDataList.size() - 1).getKey(), newlesson.getSubject(), oldDataList.get(oldDataList.size() - 1).getValue());
                                           List<Attendance> ollAttendanceList = ExelMagazinDialog.readlistGradesMagazin(cod, tempLesson);
                                           for (int q= 0; q < ollAttendanceList.size(); q++){
                                               if (ollAttendanceList.get(q).getMark_Attendance_Assessment().contains("ВИБУВ")){
                                                   ollAttendanceList.get(q).setDate(newlesson.getDate());
                                               } else{
                                                   ollAttendanceList.get(q).setDate(newlesson.getDate());
                                                   ollAttendanceList.get(q).setMark_Attendance_Assessment(" ");
                                               }
                                           }
                                           submenuMagazine(scanner, cod, newlesson, newgroup, ollAttendanceList, lessOld_or_New);
                                       }
                                   } else {
                                        //змінився
                                           List<Attendance> ollAttendanceList = ExelMagazinDialog.updateMagazin(newgroup, newlesson, cod);
                                           Group newGroup = ExelMagazinDialog.readGroup(cod);
                                           submenuMagazine(scanner, cod, newlesson, newGroup, ollAttendanceList, lessOld_or_New);
                                   }
                               }

                           } else {//коли лист EXEL НЕ існує
                               //Створюємо сутність класу Заняття
                               List<SimpleEntry<String, String>> oldDataList = new ArrayList<>();
                               Lesson newlesson = new Lesson (newsubject.subject, oldDataList);

                               ExelMagazinDialog.newExelMagazine (cod, newgroup, newsubject, newlesson);
                               List<Attendance> ollAttendanceList = new ArrayList<>();
                               for (int q = 0; q < newgroup.getFullNameStudents().size();q++)
                               {
                                   Attendance attendance = new Attendance(newlesson.date, newgroup.getFullNameStudents().get(q), " ");
                                   ollAttendanceList.add(attendance);
                               }
                               submenuMagazine(scanner, cod, newlesson, newgroup, ollAttendanceList, lessOld_or_New);
                           }
                           } catch (FileNotFoundException e) {
                           throw new RuntimeException(e);
                       } catch (IOException e) {
                           throw new RuntimeException(e);
                       }
                       break;
                    case (4):
                       exit = true;
                       System.out.println("До побачення!");
                        break;
                    default:
                       System.out.println("Невірний вибір. Спробуйте ще раз.");
                        System.out.println("Натисніть будь-яку клавішу для продовження.");
                        scanner.nextLine();
                }
            }
        scanner.close();
    }

//================//Методи:

    //Очистка екрану (Працює натисканням кнопки (Очистити) на розгорнутій на весь екран консолі)
       public static void cls()  {
        Robot bot = null;
        try {
            bot = new Robot();
        } catch (AWTException e) {
            throw new RuntimeException(e);
        }
        bot.mouseMove(14, 268);
        bot.mousePress(InputEvent.BUTTON1_MASK);
        bot.mouseRelease(InputEvent.BUTTON1_MASK);
        try {
            Thread.sleep(100);
        } catch (InterruptedException e){
        e.printStackTrace();
       }
        // Варіант 1 (Не працює)
//        try {
//           if (System.getProperty("os.name").contains("Windows")) {
//               new ProcessBuilder("cmd", "/c", "cls").inheritIO().start().waitFor();
//          } else {
//           System.out.print("\033[H\033[2J");
//           System.out.flush();
//            }
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
        // Варіант 2 Встановлення системи Jansi (не працює)
//        try {
//            AnsiConsole.systemInstall();
//            System.out.print("\033[H\033[2J");
//            System.out.flush();
//            AnsiConsole.systemUninstall(); 
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
        //Варіант 3 (не працює)
//        try {
//        Terminal terminal = TerminalFacade.createTerminal();
//        terminal.clearScreen();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
           //Варіант 4 (не працює)
//           try {
//           Runtime runtime = getRuntime();
//           Process process = runtime.exec("cls");
//           } catch (Exception e) {
//            e.printStackTrace();
//       }
       }

    //метод підменю класу студент
    public static void submenuStudent(Scanner scanner)  {
        boolean back = false;
        while (!back) {
            cls();
            System.out.println(" ");
            System.out.println("ПІДМЕНЮ СТУДЕНТИ");
            System.out.println("---------------------------------------------------");
            System.out.println("1 - Додати студента");
            System.out.println("2 - Видалити студента");
            System.out.println("3 - Коригувати інформацію");
            System.out.println("4 - Вивести таблицю студентів");
            System.out.println("5 - Повернутись в головне меню");
            System.out.println("---------------------------------------------------");
            System.out.print("Введіть ваш вибір: ");
            while (!scanner.hasNextInt()) {
                System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                scanner.next();
            }
            int choice = scanner.nextInt();
            scanner.nextLine ();

            switch (choice) {
                case 1:
                    cls();
                    System.out.println("---------------------------------------------------");
                    System.out.print("Введіть фамілію студента: ");
                    String nslastName = scanner.nextLine ();
                    System.out.print("Введіть Ім'я: ");
                    String nsfirstName = scanner.nextLine ();
                    System.out.print("Введіть по батькові: ");
                    String nspatronymic = scanner.nextLine ();
                    System.out.print("Введіть номер групи студента: ");
                    String nsgroupNumber = scanner.nextLine ();
                    if (!nslastName.isEmpty() && !nsfirstName.isEmpty() && !nspatronymic.isEmpty() && !nsgroupNumber.isEmpty()) {
                        Student newstudent = new Student(nslastName, nsfirstName, nspatronymic, nsgroupNumber);
                        Student.addToStudent(newstudent);
                    }else {
                        System.out.println("\n Ви пропустили якесь з полів спробуйте ще раз.");
                        System.out.println("Натисніть будь-яку клавішу для продовження.");
                        scanner.nextLine();
                    }
                    break;
                case 2:
                    cls();
                    List<Student> students = Student.readStudent();
                    Student.printStudent(students);
                    System.out.print("Введіть номер студента в списку для його видалення: ");
                    while (!scanner.hasNextInt()) {
                        System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                        scanner.next();
                    }
                    int number = scanner.nextInt();
                    scanner.nextLine ();
                    Student.removeStudent(number);
                    break;
                case 3:
                    cls();
                    students = Student.readStudent();
                    Student.printStudent(students);
                    System.out.println("---------------------------------------------------");
                    System.out.print("Введіть номер студента данні якого треба скоригувати: ");
                    number = scanner.nextInt ();
                    scanner.nextLine ();
                    System.out.print("Введіть фамілію студента: ");
                    nslastName = scanner.nextLine ();
                    System.out.print("Введіть Ім'я: ");
                    nsfirstName = scanner.nextLine ();
                    System.out.print("Введіть по батькові: ");
                    nspatronymic = scanner.nextLine ();
                    System.out.print("Введіть номер групи студента: ");
                    nsgroupNumber = scanner.nextLine ();
                    if (!nslastName.isEmpty() && !nsfirstName.isEmpty() && !nspatronymic.isEmpty() && !nsgroupNumber.isEmpty()) {
                        Student newstudent = new Student(nslastName, nsfirstName, nspatronymic, nsgroupNumber);
                        Student.modifyStudent(number, newstudent);
                    }
                    else {
                        System.out.println("\n Поля некоректно заповнені. Спробуйте ще раз");
                        System.out.println("Натисніть будь-яку клавішу для продовження.");
                        scanner.nextLine();
                    }
                    break;
                case 4:
                    cls();
                    students = Student.readStudent();
                    Student.printStudent(students);
                    break;
                case 5:
                    cls();
                    back = true; //вийти в основне меню
                    break;
                default:
                    System.out.println("\n Невірний вибір. Спробуйте ще раз.");
                    System.out.println("Натисніть будь-яку клавішу для продовження.");
                    scanner.nextLine();
            }
        }
    }

    //метод підменю класу викладач
    public static void submenuTeacher(Scanner scanner) {
        boolean back = false;
        while (!back) {
            cls();
            System.out.println("ПІДМЕНЮ ВИКЛАДАЧІ");
            System.out.println("---------------------------------------------------");
            System.out.println("1 - Додати викладача чи предмет, який він читає");
            System.out.println("2 - Видалити викладача");
            System.out.println("3 - Коригувати інформацію");
            System.out.println("4 - Вивести таблицю викладачів та предметів");
            System.out.println("5 - Повернутись в головне меню");
            System.out.println("---------------------------------------------------");
            System.out.print("Введіть ваш вибір: ");
            System.out.print("Введіть номер студента в списку для його видалення: ");
            while (!scanner.hasNextInt()) {
                System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                scanner.next();
            }
            int choice = scanner.nextInt();
            scanner.nextLine ();

            switch (choice) {
                case 1:
                    cls();
                    System.out.println("---------------------------------------------------");
                    System.out.print("Введіть фамілію викладача: ");
                    String ntlastName = scanner.nextLine ();
                    System.out.print("Введіть ім'я викладача: ");
                    String ntfirstName = scanner.nextLine ();
                    System.out.print("Введіть по батькові викладача: ");
                    String ntpatronymic = scanner.nextLine ();
                    System.out.print("Введіть предмет, який читає викладач: ");
                    String ntsubject = scanner.nextLine ();
                    System.out.print("Введіть номер групи в якій викладач є куратором чи натисність Ввід: ");
                    String ntcuratorOfGroup = scanner.nextLine ();
                    System.out.print("Введіть номери груп в яких викладач читає цей предмет: ");
                    String ntteacherOfGroup = scanner.nextLine ();
                    if (!ntlastName.isEmpty() && !ntfirstName.isEmpty() && !ntpatronymic.isEmpty() && !ntsubject.isEmpty() && !ntteacherOfGroup.isEmpty()) {
                        if(ntcuratorOfGroup == null)ntcuratorOfGroup = "";
                        Teacher newteacher = new Teacher(ntlastName, ntfirstName, ntpatronymic, ntsubject, ntcuratorOfGroup, ntteacherOfGroup);
                        Teacher.addTeacher(newteacher);
                    }else {
                        System.out.println("\n Ви пропустили якесь з обов'язкових полів. Спробуйте ще раз.");
                        System.out.println("Натисніть будь-яку клавішу для продовження.");
                        scanner.nextLine();
                    }
                    break;
                case 2:
                    cls();
                    List <Teacher> teachers = Teacher.readTeacher();
                    Teacher.printTeacher(teachers);
                    System.out.print("Введіть номер рядка в списку для його видалення: ");
                    int number = scanner.nextInt ();
                    scanner.nextLine ();
                    Teacher.removeTeacher(number);
                    break;
                case 3:
                    cls();
                    teachers = Teacher.readTeacher();
                    Teacher.printTeacher(teachers);
                    System.out.println("---------------------------------------------------");
                    System.out.print("Введіть номер рядка в списку для коригування інформації в ньому: ");
                    number = scanner.nextInt ();
                    scanner.nextLine ();
                    System.out.print("Введіть фамілію викладача: ");
                    ntlastName = scanner.nextLine ();
                    System.out.print("Введіть ім'я викладача: ");
                    ntfirstName = scanner.nextLine ();
                    System.out.print("Введіть по батькові викладача: ");
                    ntpatronymic = scanner.nextLine ();
                    System.out.print("Введіть предмет, який читає викладач: ");
                    ntsubject = scanner.nextLine ();
                    System.out.print("Введіть номер групи в якій викладач є куратором чи натисність Ввід: ");
                     ntcuratorOfGroup = scanner.nextLine ();
                    System.out.print("Введіть номери груп в яких викладач читає цей предмет чи натисність Ввід: ");
                     ntteacherOfGroup = scanner.nextLine ();
                    if (!ntlastName.isEmpty() && !ntfirstName.isEmpty() && !ntpatronymic.isEmpty() && !ntsubject.isEmpty() && !ntteacherOfGroup.isEmpty()) {
                        if(ntcuratorOfGroup == null)ntcuratorOfGroup = "";
                        Teacher newteacher = new Teacher(ntlastName, ntfirstName, ntpatronymic, ntsubject, ntcuratorOfGroup, ntteacherOfGroup);
                        Teacher.addTeacher(newteacher);
                    }else {
                        System.out.println("\n Ви пропустили якесь з обов'язкових полів. Спробуйте ще раз.");
                        System.out.println("Натисніть будь-яку клавішу для продовження.");
                        scanner.nextLine();
                    }
                    break;
                case 4:
                    cls();
                    teachers = Teacher.readTeacher();
                    Teacher.printTeacher(teachers);
                    break;
                case 5:
                    cls();
                    back = true; //вийти в основне меню
                    break;
                default:
                    System.out.println("\n Невірний вибір. Спробуйте ще раз.");
                    System.out.println("Натисніть будь-яку клавішу для продовження.");
                    scanner.nextLine();
            }
        }
    }
//Метод підменю журналу
public static void submenuMagazine (Scanner scanner, String cod, Lesson lesson, Group group, List<Attendance> ollAttendanceList,boolean lessOld_or_New)  {

    boolean back = false;
    while (!back) {
        cls();
        System.out.println("\n Робота з журналом.");
        System.out.println("---------------------------------------------------");
        System.out.println("1 - Перевірити присутність");
        System.out.println("2 - Оцінювання студентів");
        System.out.println("3 - Вивести журнал");
        System.out.println("4 - Повернутись в головне меню");
        System.out.println("---------------------------------------------------");
        System.out.print("Введіть ваш вибір: ");
        while (!scanner.hasNextInt()) {
            System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
            scanner.next();
        }
        int choice = scanner.nextInt();
        scanner.nextLine ();

        switch (choice) {
            case 1:
                cls();
                if (!lessOld_or_New) {
                    System.out.println("Перевіримо відвідування:");
                    for (int j = 0; j < ollAttendanceList.size(); j++) {
                        System.out.println("\n----------------------------------------");
                        if (ollAttendanceList.get(j).mark_Attendance_Assessment.equals("ВИБУВ")){
                            System.out.println((1 + j) + " " + group.fullNameStudents.get(j) + " - ВИБУВ");
                            } else {
                                System.out.println((1 + j) + " " + group.fullNameStudents.get(j));
                                boolean bool = false;
                                while (!bool) {
                                    System.out.println("\n Оберіть потрібне:");
                                    System.out.println("1 - Присутній");
                                    System.out.println("2 - Відсутній");
                                    System.out.print("Введіть необхідне: ");
                                    while (!scanner.hasNextInt()) {
                                        System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                                        scanner.next();
                                    }
                                    int key = scanner.nextInt();
                                    scanner.nextLine ();
                                    switch (key) {
                                        case 1:
                                            ollAttendanceList.get(j).setMark_Attendance_Assessment(" ");
                                            bool = true;
                                            break;
                                        case 2:
                                            ollAttendanceList.get(j).setMark_Attendance_Assessment("Відсутній");
                                            bool = true;
                                            break;
                                        default:
                                        System.out.println("Ваш вибір хибний! Спробуйте ще раз.");
                                        System.out.println("Натисніть будь-яку клавішу.");
                                        scanner.nextLine();
                                    }
                                }
                            }
                        }
                    lessOld_or_New = true;
                } else {
                    System.out.println("Скоригувати раніше введену інформацію:");
                    System.out.println("_________________________________________________________________");
                    System.out.printf("| %-5s | %-35s | %-15s |\n", "№", "ПІП студента", "Присутність");
                    System.out.println("|_______|_____________________________________|_________________|");
                    for (int i = 0; i < ollAttendanceList.size(); i++) {
                        System.out.printf("| %-5d | %-35s | %-15s |\n", 1+i, ollAttendanceList.get(i).getFulNameStudent(), ollAttendanceList.get(i).getMark_Attendance_Assessment());
                    }
                    System.out.println("|_______|_____________________________________|_________________|");
                    boolean lock = false;
                    while (!lock) {
                        System.out.print("Оберіть номер студента в списку чи натисніть \"0\" для виходу: ");
                        int num = scanner.nextInt();
                        scanner.nextLine();
                        if (num > 0 && num <= ollAttendanceList.size()) {
                            boolean bool = false;
                            while (!bool) {
                                System.out.println("\n Оберіть потрібне:");
                                System.out.println("1 - Присутній");
                                System.out.println("2 - Відсутній");
                                System.out.print("Введіть необхідне: ");
                                while (!scanner.hasNextInt()) {
                                    System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                                    scanner.next();
                                }
                                int key = scanner.nextInt();
                                scanner.nextLine ();
                                if (key == 2) {
                                    ollAttendanceList.get(num - 1).setMark_Attendance_Assessment("Відсутній");
                                    bool = true;
                                } else {
                                    if (key == 1) {
                                        ollAttendanceList.get(num - 1).setMark_Attendance_Assessment(" ");
                                        bool = true;
                                    } else {
                                        System.out.println("Ваш вибір хибний! Спробуйте ще раз.");
                                        System.out.println("Натисніть будь-яку клавішу.");
                                        scanner.nextLine();
                                    }
                                }
                            }
                        } else {
                            if (num == 0) {
                                lock = true;
                            } else {
                                System.out.println("Обраний невірний номер! Спробуйте ще раз.");
                                System.out.println("Натисніть будь-яку клавішу.");
                                scanner.nextLine();
                            }
                        }
                    }
                }
                break;
            case 2:
                cls();
                if (lessOld_or_New){
                    if (lesson.getType().contains("Контрольна робота")) {
                        System.out.println("Виставте присутнім студентам оцінки за контрольну роботу:");
                        for (int i = 0; i < ollAttendanceList.size(); i++) {
                            if (ollAttendanceList.get(i).getMark_Attendance_Assessment().equals(" ")){
                                System.out.println("______________________________");
                                System.out.println ((i+1) + " " + ollAttendanceList.get(i).getFulNameStudent());
                                boolean bool = false;
                                while (!bool) {
                                    System.out.println("\n Оцініть роботу студента:");
                                    System.out.println("1- не задовільно");
                                    System.out.println("2- не задовільно");
                                    System.out.println("3- задовільно");
                                    System.out.println("4- добре");
                                    System.out.println("5- відмінно");
                                    System.out.print("Оберіть оцінку: ");
                                    while (!scanner.hasNextInt()) {
                                        System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                                        scanner.next();
                                    }
                                    int key = scanner.nextInt();
                                    scanner.nextLine ();
                                    if (key > 0 && key <=5) {
                                        ollAttendanceList.get(i).setMark_Attendance_Assessment(String.valueOf(key));
                                        bool = true;
                                    } else {
                                        System.out.println ("Оцінка не з діапазону! Спробуйте ще раз.");
                                        System.out.println("Натисніть будь-яку клавішу.");
                                        scanner.nextLine();
                                    }
                                }
                            }
                        }
                    } else {
                        System.out.println("Список студентів:");
                        System.out.println("_________________________________________________________________");
                        System.out.printf("| %-5s | %-35s | %-15s |\n", "№", "ПІП студента", "Присутність");
                        System.out.println("|_______|_____________________________________|_________________|");
                        for (int i = 0; i < ollAttendanceList.size(); i++) {
                            System.out.printf("| %-5d | %-35s | %-15s |\n", 1+i, ollAttendanceList.get(i).getFulNameStudent(), ollAttendanceList.get(i).getMark_Attendance_Assessment());
                        }
                        System.out.println("|_______|_____________________________________|_________________|");
                        boolean lock = false;
                        while (!lock) {
                            System.out.print("\n Оберіть номер студента в списку чи натисніть \"0\" для виходу: ");
                            int num = scanner.nextInt();
                            if (num > 0 && num <= ollAttendanceList.size()) {
                                boolean bool = false;
                                while (!bool) {
                                    System.out.println("Оцініть роботу студента:");
                                    System.out.println("1- не задовільно");
                                    System.out.println("2- не задовільно");
                                    System.out.println("3- задовільно");
                                    System.out.println("4- добре");
                                    System.out.println("5- відмінно");
                                    System.out.print("Оберіть оцінку: ");
                                    while (!scanner.hasNextInt()) {
                                        System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                                        scanner.next();
                                    }
                                    int key = scanner.nextInt();
                                    scanner.nextLine ();
                                    if (key > 0 && key <= 5) {
                                        if (ollAttendanceList.get(num - 1).getMark_Attendance_Assessment().contains("ВИБУВ") ||
                                                ollAttendanceList.get(num - 1).getMark_Attendance_Assessment().contains("Відсутній")){
                                            System.out.println("Студент відсутній на занятті. Оберіть іншого");
                                            bool = true;
                                        } else {
                                        ollAttendanceList.get(num - 1).setMark_Attendance_Assessment(String.valueOf(key));
                                        bool = true;
                                        }
                                    } else {
                                        System.out.println("Оцінка не з діапазону! Спробуйте ще раз.");
                                        System.out.println("Натисніть будь-яку клавішу.");
                                        scanner.nextLine();
                                    }
                                }
                            } else {
                                if (num == 0) {
                                    lock = true;
                                } else{
                                System.out.println("Обраний невірний номер! Спробуйте ще раз.");
                                System.out.println("Натисніть будь-яку клавішу.");
                                scanner.nextLine();
                                }
                            }
                        }
                    }
                } else{
                    System.out.println ("Почніть з перевірки відвідування");
                    System.out.println("Натисніть будь-яку клавішу.");
                    scanner.nextLine();
                }
                break;
            case 3:
                    cls();
                    ExelMagazinDialog.contributeExelList(cod, lesson, ollAttendanceList);
                    ExelMagazinDialog.printExelList(cod);
                    System.out.println("Відкрити відповідну книгу EXEL для перегляду журналу? Y-так/N-ні");
                    String sw = scanner.nextLine();
                    if (sw.equals("Y") || sw.equals("y")){
                        ExelMagazinDialog.openFile();
                        System.out.println("НЕ ЗАБУДЬТЕ ЗАКРИТИ ФАЙЛ І РЕДАКТОР EXEL!");
                        System.out.println("Натисніть будь-яку клавішу для продовження роботи.");
                        scanner.nextLine();
                    }
                break;
            case 4:
                back = true;
                break;
            default:
                System.out.println("\n Невірний вибір. Спробуйте ще раз.");
                System.out.println("Натисніть будь-яку клавішу для продовження.");
                scanner.nextLine();
            }
        }
            ExelMagazinDialog.contributeExelList(cod, lesson, ollAttendanceList);
    }
}
