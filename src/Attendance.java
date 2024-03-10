// Клас, який представляє відвідування студентом певного заняття
public class Attendance {
    String date;        // Дата відвідування
    String fulNameStudent;  // Студент, який відвідав заняття ФІО
    String mark_Attendance_Assessment;    // маркер (присутності чи оцінка)


    // Конструктор класу Attendance
    public Attendance(String date, String fulNameStudent, String mark_Attendance_Assessment) {
        this.date = date;
        this.fulNameStudent = fulNameStudent;
        this.mark_Attendance_Assessment = mark_Attendance_Assessment;
    }

    //Методи
    public String getDate() {
        return date;
    }

    public String getFulNameStudent() {
        return fulNameStudent;
    }
    public String getMark_Attendance_Assessment() {
        return mark_Attendance_Assessment;
    }
    public void setMark_Attendance_Assessment(String mark_Attendance_Assessment) {
        this.mark_Attendance_Assessment = mark_Attendance_Assessment;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public void setFulNameStudent(String fulNameStudent) {
        this.fulNameStudent = fulNameStudent;
    }
}
