import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.*;
import java.io.*;
import java.util.*;
import java.util.AbstractMap.SimpleEntry;
import java.util.List;

public class ExelMagazinDialog {
    //Метод створення журналу
    public static void newExelMagazine (String cod, Group group, Subject subject, Lesson lessons) {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.createSheet(cod);
            //переведення всіх комірок в текстовий формат
            for (Row row : sheet) {
                for (Cell cell : row) {
                    CellStyle textStyle = workbook.createCellStyle();
                    DataFormat format = workbook.createDataFormat();
                    textStyle.setDataFormat(format.getFormat("@"));
                    cell.setCellStyle(textStyle);
                }
            }
            sheet.setColumnWidth(0, 4200);
            sheet.setColumnWidth(1, 14000);
            sheet.setColumnWidth(2, 5200);
            CellStyle styleBoldCenter = ExcelStyleUtils.createStyle(workbook, true, 0, HorizontalAlignment.CENTER);
            CellStyle styleLeft = ExcelStyleUtils.createStyle(workbook, false, 1, HorizontalAlignment.LEFT);
            CellStyle styleRight = ExcelStyleUtils.createStyle(workbook, true, 0, HorizontalAlignment.RIGHT);
            CellStyle styleTableBoldCenter = ExcelStyleUtils.createStyle(workbook, true, 2, HorizontalAlignment.CENTER);
            CellStyle styleTableCenter = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.CENTER);
            CellStyle styleTableLeft = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.LEFT);
            Row row = sheet.createRow(0);
            row.createCell(1).setCellValue("Журнал контролю відвідування та успішності");
            row.getCell(1).setCellStyle(styleBoldCenter);
            row = sheet.createRow(1);
            row.createCell(0).setCellValue("Предмет:");
            row.getCell(0).setCellStyle(styleRight);
            row.createCell(1).setCellValue(subject.getSubject());
            row.getCell(1).setCellStyle(styleLeft);
            row = sheet.createRow(2);
            row.createCell(0).setCellValue("Викладач:");
            row.getCell(0).setCellStyle(styleRight);
            row.createCell(1).setCellValue(subject.getFulNameteacher());
            row.getCell(1).setCellStyle(styleLeft);
            row = sheet.createRow(3);
            row.createCell(0).setCellValue("Група №:");
            row.getCell(0).setCellStyle(styleRight);
            row.createCell(1).setCellValue(group.getGroupNumber());
            row.getCell(1).setCellStyle(styleLeft);
            row = sheet.createRow(4);
            row.createCell(0).setCellValue("Куратор групи:");
            row.getCell(0).setCellStyle(styleRight);
            row.createCell(1).setCellValue(group.getFullNamesCurator());
            row.getCell(1).setCellStyle(styleLeft);

            sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 0));
            sheet.addMergedRegion(new CellRangeAddress(6, 7, 1, 1));
            row = sheet.createRow(6);
            row.createCell(0).setCellValue(" № з/п");
            row.getCell(0).setCellStyle(styleTableBoldCenter);
            row.createCell(1).setCellValue(" ПІП студента");
            row.getCell(1).setCellStyle(styleTableBoldCenter);
            row.createCell(2).setCellValue(lessons.getType());
            row.getCell(2).setCellStyle(styleTableBoldCenter);
            row = sheet.createRow(7);
            row.createCell(2).setCellValue(lessons.getDate());
            row.getCell(2).setCellStyle(styleTableBoldCenter);

            for (int n = 0; n < group.getFullNameStudents().size(); n++) {
                    row = sheet.createRow(8 + n);
                    row.createCell(0).setCellValue(1 + n);
                    row.getCell(0).setCellStyle(styleTableCenter);
                    row.createCell(1).setCellValue(group.getFullNameStudents().get(n));
                    row.getCell(1).setCellStyle(styleTableLeft);
            }
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

    // Метод коригування списку студентів в журналі
    public static List<Attendance> updateMagazin (Group newGroup, Lesson lesson, String cod){
        List<Attendance> attendanceList = new ArrayList<>();
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(cod);
            List<String> studList = new ArrayList<>();
            CellStyle styleTableCenter = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.CENTER);
            CellStyle styleTableLeft = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.LEFT);
            //Формуємо в attendanceList
            for (int i=8; i<= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
               studList.add(row.getCell(1).getStringCellValue());
                if (newGroup.getFullNameStudents().contains(row.getCell(1).getStringCellValue())){
                    attendanceList.add(new Attendance(lesson.getDate(), row.getCell(1).getStringCellValue(), " "));
                } else {
                    attendanceList.add(new Attendance(lesson.getDate(), row.getCell(1).getStringCellValue(), "ВИБУВ"));
                }
            }
            //Формуємо оновлений список студентів (відсортований)
           for (int j=0; j < newGroup.getFullNameStudents().size(); j++) {
                if (!studList.contains(newGroup.getFullNameStudents().get(j))) {
                    attendanceList.add(new Attendance(lesson.getDate(), newGroup.getFullNameStudents().get(j), " "));
                    studList.add(newGroup.getFullNameStudents().get(j));
                }
            }
            Collections.sort(studList);
            attendanceList.sort(Comparator.comparing(Attendance::getFulNameStudent));

            //Коригуєм список студентів певної групи в листі
                for (int k = 0; k < studList.size(); k++){
                    Row row = sheet.getRow(8+k);
                    row.createCell(0).setCellValue(1+k);
                    row.getCell(0).setCellStyle(styleTableCenter);

                    if (!studList.get(k).equals(row.getCell(1).getStringCellValue())){
                        sheet.shiftRows(8+k, sheet.getLastRowNum(),1);

                        Row newRow = sheet.createRow(8+k);
                        newRow.createCell(0).setCellValue(1+k);
                        newRow.getCell(0).setCellStyle(styleTableCenter);

                        newRow.createCell(1).setCellValue(studList.get(k));
                        newRow.getCell(1).setCellStyle(styleTableLeft);
                    }
                }
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
        return attendanceList;
    }
    //Метод считування класу груп з EXEL
    public static Group readGroup (String cod) {

        FileInputStream fis = null;
        XSSFWorkbook workbook = null;
        Group oldGroup = null;
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);

            Sheet sheetGroup = workbook.getSheet(cod);
            List <String> oldFulNameStudent = new ArrayList<>();
            Row Strow = sheetGroup.getRow(3);
            String groupNum = Strow.getCell(1).getStringCellValue();
            Strow = sheetGroup.getRow(4);
            String NamesCurator = Strow.getCell(1).getStringCellValue();
            for (int i = 8; i <= sheetGroup.getLastRowNum(); i++) {
                Strow = sheetGroup.getRow(i);
                String fulNameStud = Strow.getCell(1).getStringCellValue();
                oldFulNameStudent.add(fulNameStud);
                oldFulNameStudent.stream().sorted(Comparator.naturalOrder());
            }
            oldGroup = new Group (groupNum,oldFulNameStudent, NamesCurator);
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
        return oldGroup;
    }
    // Mетод считування з журналу дат і типу занять
    public static List<SimpleEntry<String, String>> readData (String cod) {
        List<SimpleEntry<String, String>> oldDataList = new ArrayList<>();

        try (FileInputStream file = new FileInputStream("resources/Univer.xlsx")) {
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(cod);
            Row row = sheet.getRow(6);
             for (int i = 2; i < row.getLastCellNum(); i++) {
                 row = sheet.getRow(7);
                 String cell1 = String.valueOf(row.getCell(i));
                 row = sheet.getRow(6);
                 String cell2 = String.valueOf(row.getCell(i));
                 SimpleEntry<String, String> pair = new SimpleEntry<>(cell1, cell2);
                oldDataList.add(pair);
             }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return oldDataList;
    }
    //Метод считування оцінок з журналу проведеного заняття
    public static List<Attendance> readlistGradesMagazin (String cod, Lesson lesson){
        List<SimpleEntry<String, String>> oldDataList = readData (cod);
        List<Attendance> attendanceList = new ArrayList<>();
        try (FileInputStream file = new FileInputStream("resources/Univer.xlsx")) {
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(cod);
            int poz = 0;
            for (int j=0; j < oldDataList.size(); j++) {
                if (lesson.getDate().equals(oldDataList.get(j).getKey()) && lesson.getType().equals(oldDataList.get(j).getValue())) {
                    poz = j+2;
                    break;
                }
            }
            Row row = sheet.getRow(7);
            String cellData = String.valueOf(row.getCell(poz));
            for (int i = 8; i <= sheet.getLastRowNum(); i++) {
                row = sheet.getRow(i);
                String cellFNS = String.valueOf(row.getCell(1));
                String cellMark = String.valueOf(row.getCell(poz));
                Attendance attend = new Attendance(cellData, cellFNS, cellMark);
                attendanceList.add (attend);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return attendanceList;
    }

    //Метод дозапису оцінок до існуючого журналу
    public static void contributeExelList (String cod, Lesson lesson, List<Attendance> attendanceList){
        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;
        List<SimpleEntry<String, String>> oldDataList = ExelMagazinDialog.readData (cod);
        try {
            fis = new FileInputStream("resources/Univer.xlsx");
            workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(cod);
            CellStyle styleTableBoldCenter = ExcelStyleUtils.createStyle(workbook, true, 2, HorizontalAlignment.CENTER);
            CellStyle styleTableCenter = ExcelStyleUtils.createStyle(workbook, false, 2, HorizontalAlignment.CENTER);

            int pozCell;
            if (attendanceList !=null && lesson !=null && !oldDataList.isEmpty()){
                //визначаємо номер стовпця для запису
                pozCell = oldDataList.size()+2;
                boolean bool = true;
                for (int i = 0; i < oldDataList.size(); i++) {
                    if (lesson.getDate().equals(oldDataList.get(i).getKey()) && lesson.getType().equals(oldDataList.get(i).getValue())) {
                        pozCell = i+2; bool = false;
                        break;
                    }
                }
                //заповнюємо заголовок в разі додавання стовпця
                sheet.setColumnWidth(pozCell, 5200);
                if (bool) {
                    Row row = sheet.getRow(6);
                    row.createCell(pozCell).setCellValue(lesson.getType());
                    row.getCell(pozCell).setCellStyle(styleTableBoldCenter);
                    row = sheet.getRow(7);
                    row.createCell(pozCell).setCellValue(lesson.getDate());
                    row.getCell(pozCell).setCellStyle(styleTableBoldCenter);
                }
                 //заповнюємо стовпець даними перевіряючи чи належать вони цьому студенту
                for (int j=0; j < attendanceList.size();j++){
                    Row row = sheet.getRow(8+j);
                    if (row.getCell(1).getStringCellValue().equals(attendanceList.get(j).fulNameStudent)){
                        row.createCell(pozCell).setCellValue(attendanceList.get(j).mark_Attendance_Assessment);
                        row.getCell(pozCell).setCellStyle(styleTableCenter);
                    }
                }
            }else {
                System.out.println("До журналу не внесено жодних змін!");
            }
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

    //Виведення журналу в консоль
    public static void printExelList (String cod){
        University.cls();
        try (FileInputStream file = new FileInputStream("resources/Univer.xlsx")) {
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(cod);
            if (sheet == null) {
                System.out.println("Лист з назвою " + cod + " не знайдений.");
                workbook.close();
                return;
            }
            String str1;
            String str2;
            String str3;
            for (int j = 0; j < 5; j++) {
                Row row = sheet.getRow(j);
                if (row.getCell(0) != null) {
                    str1 = String.valueOf(row.getCell(0));
                } else {
                    str1 = " ";
                }
                if (row.getCell(1) != null) {
                    str2 = String.valueOf(row.getCell(1));
                } else {
                    str2 = " ";
                }
                System.out.println("\t" + str1 + " " + str2);
            }
            Row row = sheet.getRow(8);
            for (int z=0; z<= 45+(row.getLastCellNum()-2)*20; z++) {
                System.out.print("_");
            }
            System.out.println("_");
                for (int j = 6; j <= sheet.getLastRowNum(); j++) {
                     row = sheet.getRow(j);
                    if (row.getCell(0) != null) {
                        str1 = String.valueOf(row.getCell(0));
                    } else {
                        str1 = " ";
                    }
                    if (row.getCell(1) != null) {
                        str2 = String.valueOf(row.getCell(1));
                    } else {
                        str2 = " ";
                    }
                    if (j<8){
                        System.out.printf("| %-6s | %-35s |", str1, str2);
                    } else{
                        System.out.printf("| %-6d | %-35s |", j-7, str2);
                    }
                    for (int i = 2; i < row.getLastCellNum(); i++) {
                        if (row.getCell(i) != null) {
                            str3 = String.valueOf(row.getCell(i));
                        } else {
                            str3 = " ";
                        }
                        System.out.printf(" %-17s |", str3);
                    }
                    System.out.println();
                }
                for (int z=0; z<= 45+(row.getLastCellNum()-2)*20; z++) {
                System.out.print("_");
            }
            System.out.println("_");
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //Метод відкриття файлу EXEL в редакторі Microsoft Ofice
    public static void openFile () {
        try {
            File excelFile = new File("resources/Univer.xlsx");
            Desktop.getDesktop().open(excelFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

