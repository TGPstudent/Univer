import java.util.List;
import java.util.Optional;
import java.util.Scanner;
import java.util.stream.Collectors;

public class Subject {

    String fulNameteacher;
    String subject;
    String groupNumber;

    // Конструктор класу предмет
    public Subject (String fulNameteacher, String subject, String groupNumber) {
        this.fulNameteacher = fulNameteacher;
        this.subject  = subject;
        this.groupNumber = groupNumber;
    }
    // Конструктор класу предмет (діалог)
    public Subject (Scanner scanner, String groupNumber) {
        List<Teacher> teacherList = Teacher.readTeacher();

        List<String> menuSubject = teacherList.stream()
                .filter(teacher -> teacher.getTeacherOfGroup().contains(groupNumber))
                .map(teacher -> teacher.getSubject())
                .sorted()
                .collect(Collectors.toList());

        String selectedSub;
        if (menuSubject.size() < 2 && !menuSubject.isEmpty()){
            selectedSub = menuSubject.get(0);
            System.out.println("Предмет: " + selectedSub);
        } else {
            System.out.println("\n Оберіть предмет:");
            for (int i = 0; i < menuSubject.size(); i++) {
                System.out.println((i + 1) + ". " + menuSubject.get(i));
            }
            System.out.print("Введіть ваш вибір: ");
            int choice = scanner.nextInt();
            scanner.nextLine();
            selectedSub = menuSubject.get(choice - 1);
        }
        String selectedSubject = selectedSub;
        Optional<Teacher> teacherFullName = teacherList.stream()
                .filter(teacher -> teacher.getSubject().equals(selectedSubject) && teacher.getTeacherOfGroup().contains(groupNumber))
                .findFirst();

        if (teacherFullName.isPresent()) {
            Teacher teacher = teacherFullName.get();
            this.fulNameteacher = teacher.getFullName();
        } else {
            System.out.println("Не знайдено викладача, який читає предмет");
            System.out.print("Введіть свою фамілію, ім'я та по батькові: ");
            this.fulNameteacher = scanner.nextLine ();
        }
        this.subject  = selectedSubject;
        this.groupNumber = groupNumber;
    }

    public String getFulNameteacher() {
        return fulNameteacher;
    }

    public String getSubject() {
        return subject;
    }

    public String getGroupNumber() {
        return groupNumber;
    }
}

