import java.util.*;
import java.util.stream.Collectors;

public class Group {
    String groupNumber;              // Номер групи
    List<String> fullNameStudents;   // Список студентів у групі (ПІП)
    String fullNamesCurator;         // Викладач-куратор групи (ПІП)

    // Конструктор класу Group, який ініціалізує номер групи, список студентів та викладача-куратора
    public Group(String groupNumber, List<String> fullNameStudents, String fullNamesCurator) {
        this.groupNumber = groupNumber;
        this.fullNameStudents = fullNameStudents;
        this.fullNamesCurator = fullNamesCurator;
    }
    //Конструктор без параметрів
    public Group() {
        List<Student> studentList = Student.readStudent();
        Scanner scanner = new Scanner(System.in);

        List<String> menuGroupNumbers = studentList.stream()
                .map(Student::getGroupNumber)
                .distinct()
                .sorted()
                .collect(Collectors.toList());

        String selectedGroupNumber;
        if (menuGroupNumbers.size() < 2 && !menuGroupNumbers.isEmpty()){
            selectedGroupNumber = menuGroupNumbers.get(0);
            System.out.println("Номер групи: " + selectedGroupNumber);
        } else {
            System.out.println("Оберіть номер групи:");
            for (int i = 0; i < menuGroupNumbers.size(); i++) {
                System.out.println((i + 1) + ". " + menuGroupNumbers.get(i));
            }
            System.out.print("Введіть ваш вибір: ");
            int choice = scanner.nextInt();
            selectedGroupNumber = menuGroupNumbers.get(choice - 1);
        }

        List<String> studentFullNames = studentList.stream()
                .filter(student -> student.getGroupNumber().equals(selectedGroupNumber))
                .map(Student::getFullName)
                .sorted()
                .collect(Collectors.toList());
        
        this.groupNumber = selectedGroupNumber;
        this.fullNameStudents = studentFullNames;

        List<Teacher> teachersList = Teacher.readTeacher();
        Optional<Teacher> curatorOptional = teachersList.stream()
                .filter(teacher -> teacher.getCuratorOfGroup().equals(selectedGroupNumber))
                .findFirst();

        if (curatorOptional.isPresent()) {
            Teacher curator = curatorOptional.get();
            this.fullNamesCurator = curator.getFullName();
        } else {
            this.fullNamesCurator = "Жоден викладач не є куратором цієї групи";
        }
    }
    public List<String> getFullNameStudents() {
        return fullNameStudents;
    }
    public String getGroupNumber() {return groupNumber;}

    public String getFullNamesCurator() {return fullNamesCurator;}

}