import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.ArrayList;
import java.util.List;
import java.io.*;

public class Student extends Person {
    public String getGroupNumber() {
        return groupNumber;
    }

    String groupNumber;
//Конструктор класу
    public Student(String lastName, String firstName, String patronymic, String groupNumber) {
        super(lastName, firstName, patronymic);
        this.groupNumber = groupNumber;
    }

    // Перевизначений метод для отримання ФІO студента
    @Override
    public String getFullName() {
        return  lastName + " " + firstName + " " + patronymic;
    }


    // Метод считування даних студентів з таблиці Excel
    public static List<Student> readStudent() {

        FileInputStream fis = null;
        XSSFWorkbook workbook = null;
        List<Student> studentList = null;
        try {
            studentList = new ArrayList<>();
            fis = new FileInputStream(new File("resources/Univer.xlsx"));
            workbook = new XSSFWorkbook(fis);

            Sheet sheetStudent = workbook.getSheet("Student");
            for (int i = 1; i <= sheetStudent.getLastRowNum(); i++) {
                Row Strow = sheetStudent.getRow(i);
                String lastNameSt = Strow.getCell(0).getStringCellValue();
                String firstNameSt = Strow.getCell(1).getStringCellValue();
                String patronymicSt = Strow.getCell(2).getStringCellValue();
                String group = Strow.getCell(3).getStringCellValue();
                // Екземпляр класу Student
                Student student = new Student(lastNameSt, firstNameSt, patronymicSt, group);
                studentList.add(student);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }  finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
                if (fis != null) {
                    fis.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return studentList;
    }

    // Метод для додавання студента до таблиці Excel
    public static void addToStudent(Student student) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheetStudent = workbook.getSheet("Student");
            CellStyle styleLeft = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.LEFT);
            CellStyle styleCenter = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.CENTER);
            int lastRowNum = sheetStudent.getLastRowNum();
            Row row = sheetStudent.createRow(lastRowNum + 1);
            row.createCell(0).setCellValue(student.lastName);
                row.getCell(0).setCellStyle(styleLeft);
            row.createCell(1).setCellValue(student.firstName);
                row.getCell(1).setCellStyle(styleLeft);
            row.createCell(2).setCellValue(student.patronymic);
                row.getCell(2).setCellStyle(styleLeft);
            row.createCell(3).setCellValue(student.groupNumber);
                row.getCell(3).setCellStyle(styleCenter);
            fos = new FileOutputStream("resources/Univer.xlsx");
            workbook.write(fos);
            System.out.println("Студента було додано до списку.");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
                if (fis != null) {
                    fis.close();
                }
                if (fos != null) {
                    fos.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    // Метод видалення студента з таблиці EXEL
    public static void removeStudent(int rowIndex) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheetStudent = workbook.getSheet("Student");

            if (rowIndex < 1 || rowIndex > sheetStudent.getLastRowNum()) {
                System.out.println("Помилка: Некоректний номер рядка для видалення.");
                return;
            }
            Row row = sheetStudent.getRow(rowIndex);
            sheetStudent.removeRow(row);
            int lastRowNum = sheetStudent.getLastRowNum();

            // Зсуваємо усі рядки, які знаходяться під видаленим рядком, на один вгору
            sheetStudent.shiftRows(rowIndex+1, lastRowNum, -1);

            System.out.println("Рядок з номером " + (rowIndex-1) + " було видалено зі списку.");

            fos = new FileOutputStream("resources/Univer.xlsx");
            workbook.write(fos);

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
                if (fis != null) {
                    fis.close();
                }
                if (fos != null) {
                    fos.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    // Метод для внесення змін до даних студента в таблиці Excel
    public static void modifyStudent(int rowIndex, Student newStudent) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheetStudent = workbook.getSheet("Student");
            if (rowIndex < 1 || rowIndex > sheetStudent.getLastRowNum()) {
                System.out.println("Помилка: Некоректний номер рядка для коригування.");
            return;
            }
                Row row = sheetStudent.getRow(rowIndex);
                if (row == null) {
                    System.out.println("Помилка: Рядок для видалення не існує.");
                    return;
                }
                row.getCell(0).setCellValue(newStudent.lastName);
                row.getCell(1).setCellValue(newStudent.firstName);
                row.getCell(2).setCellValue(newStudent.patronymic);
                row.getCell(3).setCellValue(newStudent.groupNumber);

        System.out.println("Рядок з номером " + rowIndex + " було оновлено в таблиці Excel.");

        // Збереження змін у файл Excel
                fos = new FileOutputStream("resources/Univer.xlsx");
                workbook.write(fos);
        }
        catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
                if (fis != null) {
                    fis.close();
                }
                if (fos != null) {
                    fos.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    // Метод для виведення таблиці студентів в консоль
    public static void printStudent(List<Student> students) {
        int n = 1;
        System.out.println("Список студентів всіх груп");
        System.out.println("____________________________________________________________");
        System.out.printf("| %-5s | %-35s | %-10s |\n", "№", "ПІП студента", "Група №");
        System.out.println("|_______|_____________________________________|____________|");
        for (Student student : students) {
            System.out.printf("| %-5s | %-35s | %-10s |\n", n, student.getFullName(), student.groupNumber);
            n++;
        }
        System.out.println("|_______|_____________________________________|____________|");
    }
}
