package org.example;


import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.IOException;

public class util {
    public static Font getCellStyle(Cell cell) throws DocumentException, IOException {
        //System.out.println("NOT NULL: "+ cell.getRowIndex() +"," +cell.getColumnIndex());
        Font font = new Font();
        CellStyle cellStyle = cell.getCellStyle();
        org.apache.poi.ss.usermodel.Font cellFont = cell.getSheet()
                .getWorkbook()
                .getFontAt(cellStyle.getFontIndex());


        short fontColorIndex = cellFont.getColor();
        if (fontColorIndex != IndexedColors.AUTOMATIC.getIndex() && cellFont instanceof XSSFFont) {
            XSSFColor fontColor = ((XSSFFont) cellFont).getXSSFColor();
            if (fontColor != null) {
                byte[] rgb = fontColor.getRGB();
                if (rgb != null && rgb.length == 3) {
                    // System.out.println((rgb[0] & 0xFF) + " " + (rgb[1] & 0xFF) + " " + (rgb[2] & 0xFF));
                    font.setColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                }
            }
        }


        if (cellFont.getItalic()) {
            font.setStyle(Font.ITALIC);
        }

        if (cellFont.getStrikeout()) {
            font.setStyle(Font.STRIKETHRU);
        }

        if (cellFont.getUnderline() == 1) {
            font.setStyle(Font.UNDERLINE);
        }

        short fontSize = cellFont.getFontHeightInPoints();
        font.setSize(fontSize);

        if (cellFont.getBold()) {
            font.setStyle(Font.BOLD);
        }

        String fontName = cellFont.getFontName();
        if (FontFactory.isRegistered(fontName)) {
            font.setFamily(fontName); // Use extracted font family if supported by iText
        } else {
            //logger.warn("Unsupported font type: {}", fontName);
            // - Use a fallback font (e.g., Helvetica)
            font.setFamily("Helvetica");
        }

        return font;
    }
}
