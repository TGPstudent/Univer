import java.text.ParseException;
import java.util.*;
import java.text.SimpleDateFormat;

// Клас, який представляє заняття (лекцію, семінар, контрольну роботу) з певного предмету
public class Lesson {
    String date;        // Дата заняття
    String subject; // Предмет, до якого відноситься заняття
    String type; //Тип заняття (Лекція, семінар, контрольна робота)


    public Lesson (String date, String subject, String type){
        this.date =date;
        this.subject = subject;
        this.type = type;
    }

    // Конструктор класу Lesson, який ініціалізує тип заняття
    public Lesson (String subject, List<AbstractMap.SimpleEntry<String, String>> oldDataList) {
        this.subject = subject;
        this.type = "";
        Scanner scanner = new Scanner(System.in);
        boolean less_exit = false;
        while (!less_exit) {
            System.out.println("\n Оберіть метод задання дати:");
            System.out.println("1 - Пoточна дата");
            System.out.println("2 - Дата попередньо проведеного заняття");
            System.out.println("3 - Введіть дату в форматі (дд.мм.рррр.)");
            System.out.print("Введіть обраний пункт: ");
            int Gen_num = scanner.nextInt();
            scanner.nextLine();
            switch (Gen_num) {
                case 1:
                SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
                this.date = dateFormat.format(new Date());
                less_exit = true;
                System.out.println("Ок");
                break;
                case 2:
                    if (!oldDataList.isEmpty()) {
                        if (oldDataList.size() < 2) {
                            this.date = oldDataList.get(0).getKey();
                            this.type = oldDataList.get(0).getValue();
                            System.out.println("Дата попереднього заняття: " + oldDataList.get(0).getKey());
                            System.out.println("і це: " + oldDataList.get(0).getValue());
                            System.out.println("Натисни будь-яку кнопку для продовження");
                            scanner.nextLine();
                        } else {
                            boolean k = false;
                            while (!k) {
                                System.out.println("\n Оберіть потрібну дату:");
                                for (int i = 0; i < oldDataList.size(); i++) {
                                    System.out.println((i + 1) + ". " + oldDataList.get(i).getKey());
                                }
                                System.out.print("Введіть ваш вибір: ");
                                int choiceData = scanner.nextInt();
                                scanner.nextLine();
                                if (choiceData > 0 && choiceData <= oldDataList.size()) {
                                    this.date = oldDataList.get(choiceData - 1).getKey();
                                    System.out.println("В обрану дату проводили " + oldDataList.get(choiceData - 1).getValue());
                                    System.out.println("Натисни будь-яку кнопку для продовження");
                                    scanner.nextLine();
                                    this.type = oldDataList.get(choiceData - 1).getValue();
                                    k = true;
                                } else {
                                    System.out.println("Ви обрали номер дати якої нема в списку?");
                                    System.out.println("Спробуйте ще раз. Для продовження натисни будь-яку кнопку");
                                    scanner.nextLine();
                                }
                            }
                        }
                        less_exit = true;
                    } else {
                        System.out.println("Заняття з вашого предмету ще не проводились. Оберіть інший варіант задання дати");
                        System.out.println("Натисніть будь-яку клавішу для продовження.");
                        scanner.nextLine();
                    }
                    break;
                case 3:
                    System.out.print("Введіть дату \"дд.мм.рррр\": ");
                    String selectData = scanner.nextLine();
                    SimpleDateFormat dFormat = new SimpleDateFormat("dd.MM.yyyy");
                    try {
                       dFormat.parse(selectData);
                        this.date = selectData;
                        less_exit = true;
                        System.out.println("Ок");
                    } catch (ParseException e) {
                        System.out.println("Некоректний формат дати. Будь ласка, введіть дату у форматі \"дд.мм.рррр\"");
                        System.out.println("Натисніть будь-яку клавішу для продовження.");
                        scanner.nextLine();
                    }
                    break;
                default:
                    System.out.print("Ви не коректно обрали тип введення дати. Спробуйте ще раз.");
                    System.out.println("Натисніть будь-яку клавішу для продовження.");
                    scanner.nextLine();
            }
        }
        if (type !=null && type.equals("")){
            boolean l_exit  = false;
            while (!l_exit) {
                System.out.println("Оберіть тип заняття:");
                System.out.println("1 - Лекція");
                System.out.println("2 - Семінар");
                System.out.println("3 - Контрольна робота");
                System.out.print("Введіть обране: ");
                while (!scanner.hasNextInt()) {
                    System.out.println("Введене не є числом. Будь ласка, введіть число зі списку:");
                    scanner.next();
                }
                int Gen_num = scanner.nextInt();
                scanner.nextLine ();

                switch (Gen_num) {
                    case 1:
                        this.type = "Лекція";
                        l_exit = true;
                        System.out.println("Ок");
                        break;
                    case 2:
                        this.type = "Семінар";
                        l_exit = true;
                        System.out.println("Ок");
                        break;
                    case 3:
                        this.type = "Контрольна робота";
                        l_exit = true;
                        System.out.println("Ок");
                        break;
                    default:
                        System.out.print("Ви не коректно обрали тип заняття. Спробуйте ще раз.");
                        System.out.println("Натисніть будь-яку клавішу для продовження.");
                        scanner.nextLine();
                }
            }
        }
    }
    public String getDate() {
        return date;
    }
    public String getType() {
        return type;
    }
    public String getSubject() {return subject; }
}
