import java.io.*;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Teacher extends Person {
    String subject;         // Предмет, який читає викладач
    String curatorOfGroup;     // Номер групи, куратором якої є викладач
    String  teacherOfGroup;  // Номери груп, в яких читає викладач

    public Teacher(String lastName, String firstName, String patronymic,
                   String subject, String curatorOfGroup, String  teacherOfGroup) {
        super(lastName,firstName, patronymic);
        this.subject = subject;
        this.curatorOfGroup = curatorOfGroup;
        this.teacherOfGroup = teacherOfGroup;
    }
    public String getTeacherOfGroup() {
        return teacherOfGroup;
    }
    public String getCuratorOfGroup() {
        return curatorOfGroup;
    }
    public String getSubject() {
        return subject;
    }
    // Перевизначений метод для отримання ФІО викладача
    @Override
    public String getFullName() {
        return  lastName + " " + firstName + " " + patronymic;
    }

    // Метод генерування абревіатури предмету
    public static String generateShortSubject(String subject) {

        String[] words = subject.split("\\s+");
        StringBuilder shortSubjectBuilder = new StringBuilder();

        for (String word : words) {
            if (!word.isEmpty()) {
                shortSubjectBuilder.append(Character.toUpperCase(word.charAt(0)));
            }
        }
        return shortSubjectBuilder.toString();
    }

    // Метод считуввання даних викладача з таблиці Excel
    public static List<Teacher>  readTeacher (){
        FileInputStream fis = null;
        XSSFWorkbook workbook = null;
        List<Teacher> teacherList = null;
        try {
            fis = new FileInputStream(new File("resources/Univer.xlsx"));
            workbook = new XSSFWorkbook(fis);
            teacherList = new ArrayList<>();
            Sheet sheetTeacher = workbook.getSheet("Teacher");

        for (int i = 1; i <= sheetTeacher.getLastRowNum(); i++) {
            Row Tearow = sheetTeacher.getRow(i);
            String lastNameTea = Tearow.getCell(0).getStringCellValue();
            String firstNameTea = Tearow.getCell(1).getStringCellValue();
            String patronymicTea = Tearow.getCell(2).getStringCellValue();

            String subject = Tearow.getCell(3).getStringCellValue();
            String curatorOfGroup = "";
            Cell curatorCell = Tearow.getCell(4);
            if (curatorCell != null) {
                curatorOfGroup = curatorCell.getStringCellValue();
            }
            String teacherOfGroup = Tearow.getCell(5).getStringCellValue();

            // Екземпляр класу Teacher
            Teacher teacher = new Teacher(lastNameTea, firstNameTea, patronymicTea,
                    subject, curatorOfGroup, teacherOfGroup);
            teacherList.add(teacher);
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
        return teacherList;
    }

    // Метод для додавання викладача до таблиці Excel
    public static void addTeacher(Teacher teacher) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheetTeacher = workbook.getSheet("Teacher");
            if (sheetTeacher == null) {
                sheetTeacher = workbook.createSheet("Teacher");
            }
        int lastRowNum = sheetTeacher.getLastRowNum();
            CellStyle styleLeft = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.LEFT);
            CellStyle styleCenter = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.CENTER);
        Row row = sheetTeacher.createRow(lastRowNum + 1);
        row.createCell(0).setCellValue(teacher.lastName);
            row.getCell(0).setCellStyle(styleLeft);
        row.createCell(1).setCellValue(teacher.firstName);
            row.getCell(1).setCellStyle(styleLeft);
        row.createCell(2).setCellValue(teacher.patronymic);
            row.getCell(2).setCellStyle(styleLeft);
        row.createCell(3).setCellValue(teacher.subject);
            row.getCell(3).setCellStyle(styleLeft);
        row.createCell(4).setCellValue(teacher.curatorOfGroup);
            row.getCell(4).setCellStyle(styleCenter);
        row.createCell(5).setCellValue(teacher.teacherOfGroup);
            row.getCell(5).setCellStyle(styleCenter);
        fos = new FileOutputStream("resources/Univer.xlsx");
        workbook.write(fos);
            System.out.println("Внесену інформацію було додано до списку.");
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

    // Метод для видалення викладача з таблиці Excel
    public static void removeTeacher(int rowIndex) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheetTeacher = workbook.getSheet("Teacher");
            if (sheetTeacher == null) {
                sheetTeacher = workbook.createSheet("Teacher");
            }

            if (rowIndex < 1 || rowIndex > sheetTeacher.getLastRowNum()) {
                System.out.println("Помилка: Некоректний номер рядка для видалення.");
                return;
            }
            Row row = sheetTeacher.getRow(rowIndex);
            if (row == null) {
                System.out.println("Помилка: Рядок для видалення не існує.");
                return;
            }
            sheetTeacher.removeRow(row);

            int lastRowNum = sheetTeacher.getLastRowNum();
            sheetTeacher.shiftRows(rowIndex + 1, lastRowNum, -1); // Зсув рядків вгору на один

            System.out.println("Рядок з номером " + rowIndex + " було видалено зі списку.");

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

    // Метод для коригування даних викладача в таблиці Excel
    public static void modifyTeacher(int rowIndex, Teacher newTeacher) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheetTeacher = workbook.getSheet("Teacher");
            if (sheetTeacher == null) {
                sheetTeacher = workbook.createSheet("Teacher");
            }
            if (rowIndex < 1 || rowIndex > sheetTeacher.getLastRowNum()) {
                System.out.println("Помилка: Некоректний номер рядка для коригування.");
                return;
            }
            Row row = sheetTeacher.getRow(rowIndex);
            if (row == null) {
                System.out.println("Помилка: Рядок обраний для коригування не існує.");
                return;
            }
            row.getCell(0).setCellValue(newTeacher.lastName);
            row.getCell(1).setCellValue(newTeacher.firstName);
            row.getCell(2).setCellValue(newTeacher.patronymic);
            row.getCell(3).setCellValue(newTeacher.subject);
            row.getCell(4).setCellValue(newTeacher.curatorOfGroup);
            row.getCell(5).setCellValue(newTeacher.teacherOfGroup);

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

    // Метод для виведення таблиці викладачів в консоль
    public static void printTeacher(List<Teacher> teachers) {
        int n = 1;
        System.out.println();
        System.out.println("______________________________________________________________________________________________________________________________________________________");
        System.out.printf("| %-5s | %-35s | %-60s | %-15s | %-20s |\n", "№", "ПІП викладача", "Предмет", "Куратор групи", "Викладач груп");
        System.out.println("|_______|_____________________________________|______________________________________________________________|_________________|______________________|");
        for (Teacher teacher : teachers) {
            System.out.printf("| %-5s | %-35s | %-60s | %-15s | %-20s |\n", n, teacher.getFullName(), teacher.subject, teacher.curatorOfGroup, teacher.teacherOfGroup);
            n++;
        }
        System.out.println("|_______|_____________________________________|______________________________________________________________|_________________|______________________|");
    }
}

