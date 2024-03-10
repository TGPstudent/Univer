import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.*;

public class ExcelStyleUtils {

    public static CellStyle createStyle(Workbook workbook,  boolean bold, int borderBottom, HorizontalAlignment alignment) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Times New Roman");
        font.setFontHeightInPoints((short) 12);
        font.setBold(bold);
        style.setFont(font);
        style.setAlignment(alignment);
        switch (borderBottom){
            case 1:
            style.setBorderBottom(BorderStyle.THIN);
            break;
            case 2:
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            break;
    }
        return style;
    }
}