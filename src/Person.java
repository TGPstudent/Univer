// Абстрактний клас, який визначає основні атрибути та методи для особи (студента або викладача)

abstract class Person {
    String lastName;    // Прізвище особи
    String firstName;   // Ім'я особи
    String patronymic;  // По батькові особи

    // Конструктор класу, який ініціалізує атрибути особи
    public Person(String lastName, String firstName, String patronymic) {
        this.lastName = lastName;
        this.firstName = firstName;
        this.patronymic = patronymic;
    }

    // Абстрактний метод, який буде реалізований в підкласах (Student та Teacher)
    public abstract String getFullName();  // Метод для отримання повного імені особи
}

