package org.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Map;

import com.lowagie.text.*;
import com.lowagie.text.Font;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;


public class XSSFToPDF {

    private static final Logger logger = LogManager.getLogger(XSSFToPDF.class);
    public static float[] colsWidth;
    public static float[] rowsHeight;

    public static XSSFWorkbook readExcelFile(String excelFilePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        inputStream.close();
        return workbook;
    }

    private static Document createPDFDocument(String pdfFilePath) throws IOException, DocumentException {
        Document document = new Document();
        PdfWriter.getInstance(document, new FileOutputStream(pdfFilePath));
        document.open();
        return document;
    }

    public static void convertExcelToPDF(String excelFilePath, String pdfFilePath) throws IOException, DocumentException {
        Workbook workbook = readExcelFile(excelFilePath);
        Document document = createPDFDocument(pdfFilePath);
        int maxCol = 0;
        int maxRow = 0;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            int rowCount = sheet.getPhysicalNumberOfRows();

            maxRow = sheet.getLastRowNum();
            for (Row row : sheet) {
                short lastCell = row.getLastCellNum();
                if (lastCell - 1 > maxCol) {
                    maxCol = lastCell - 1;
                }
            }

            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch(); // I know it is ugly, actually you get the actual instance here
            for (XSSFShape shape : drawing.getShapes()) {
                if (shape instanceof XSSFPicture picture) {
                    XSSFPictureData xssfPictureData = picture.getPictureData();
                    ClientAnchor anchor = picture.getPreferredSize();
                    int row1 = anchor.getRow1();
                    int row2 = anchor.getRow2();
                    int col1 = anchor.getCol1();
                    int col2 = anchor.getCol2();
//                    System.out.println("Row1: " + row1 + " Row2: " + row2);
//                    System.out.println("Column1: " + col1 + " Column2: " + col2);
//                    System.out.println(anchor.getDx1()/EMU_PER_PIXEL + " " +  anchor.getDx2()/EMU_PER_PIXEL + " " + anchor.getDy1()/EMU_PER_PIXEL + " " + anchor.getDy2()/EMU_PER_PIXEL);
                    // Saving the file
//                    String ext = xssfPictureData.suggestFileExtension();
//                    byte[] data = xssfPictureData.getData();
//                    String filePath = "/Users/o_dung_quoc.p/Work/excel-file-java/image1.png";
//                    try (FileOutputStream os = new FileOutputStream(filePath)) {
//                        os.write(data);
//                        os.flush();
//                    }
                    if (row2 > maxRow) {
                        maxRow = row2;
                    }
                    if (col2 > maxCol) {
                        maxCol = col2;
                    }
                }
            }




//            for (int j = 0; j < 10; j++) {
//                System.out.println("Default row height: " + sheet);
//            }
//            title.setSpacingAfter(20f);
//            title.setAlignment(Element.ALIGN_CENTER);
//            document.add(title);

            initData(sheet, maxRow, maxCol);
            setupTables(sheet, document, maxRow, maxCol);
            // Add a new page for each sheet (except the last one)
            if (i < workbook.getNumberOfSheets() - 1) {
                document.newPage();
            }
        }

        document.close();
        workbook.close();
    }

    private static void setupTables(Sheet sheet, Document document, int maxRow, int maxCol) throws DocumentException, IOException {
        float margin = 36;
        int currCol = 0;

        while (currCol != maxCol) {
            float totalPixelWidth = 0;
            for (int c = currCol;; c++) {
;
                if (totalPixelWidth + colsWidth[c] > document.getPageSize().getWidth() || c == maxCol) {
                    if (c == maxRow) {
                        createTable(sheet, document, maxRow, currCol, c , totalPixelWidth);
                    } else {
                        createTable(sheet, document, maxRow, currCol, c - 1 , totalPixelWidth);
                    }

                    currCol = c;
                    break;
                } else{
                    totalPixelWidth += colsWidth[c];
                }
            }
        }
    }

    private static void initData(Sheet sheet, int maxRow, int maxCol) {
        colsWidth = new float[maxCol+1];
        rowsHeight = new float[maxRow+1];
        for (int i = 0; i <= maxCol;i++) {
            colsWidth[i] = sheet.getColumnWidthInPixels(i);
        }

        for (int i = 0; i <= maxRow;i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                rowsHeight[i] = sheet.getDefaultRowHeightInPoints();
            } else {
                rowsHeight[i] = row.getHeightInPoints();
            }
        }
    }


    public static String getCellText(Cell cell, CellType type) {
        String cellValue = "";
        if (cell == null) {
            return "";
        }
        switch (type) {
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case NUMERIC:
                double value = cell.getNumericCellValue();
                if (Math.floor(value) == value || Math.ceil(value) == value) {
                    cellValue = String.valueOf((long) value);
                } else {
                    cellValue = String.valueOf(value);
                }

                break;
            case FORMULA:
                return getCellText(cell, cell.getCachedFormulaResultType());
            default:
                cellValue = "";
                break;
        }
        return cellValue;
    }

    private static void createTable(Sheet sheet, Document document, int maxRow, int startCol, int endCol, float totalWidth) throws DocumentException, IOException {

        PdfPTable table = new PdfPTable(1);
        //table.setWidthPercentage(totalWidth/595);
        table.setTotalWidth(totalWidth);
        table.setHorizontalAlignment(Element.ALIGN_LEFT);
        //table.
        table.getDefaultCell().setBorder(Rectangle.NO_BORDER);
        System.out.println(0+" "+maxRow +" "+startCol+" "+ endCol);
        for (int currRow = 0; currRow <= maxRow ; currRow++) {
            Row row = sheet.getRow(currRow);
            if (row == null) {
                PdfPCell cell = new PdfPCell(new Phrase("", null));
                cell.setBorder(Rectangle.NO_BORDER);
                cell.setFixedHeight(rowsHeight[currRow]);
                table.addCell(cell);
            } else {

                PdfPTable nested = new PdfPTable(endCol - startCol + 1);

                //Set width

                float[] widths = Arrays.copyOfRange(colsWidth, startCol, endCol+1);
                nested.setWidths(widths);
                //int currIndex = 0;
                for (int currCol = startCol; currCol <= endCol; currCol++) {
                    Cell cell = row.getCell(currCol);

                    //System.out.println(row.getRowNum()+ ","+ currIndex);
                    PdfPCell cellPdf = null;
                    if (cell != null) {
                        CellType type = cell.getCellType();
                        String cellValue = getCellText(cell, type);
                        cellPdf = new PdfPCell(new Phrase(cellValue, getCellStyle(cell)));
                        setBackgroundColor(cell, cellPdf);
                        setCellAlignment(cell, cellPdf);

                    } else {
                        cellPdf = new PdfPCell(new Phrase("", null));
                    }
                    cellPdf.setFixedHeight(rowsHeight[currRow]);
                    cellPdf.setBorder(Rectangle.NO_BORDER);
                    nested.addCell(cellPdf);

                }
                table.addCell(nested);

            }
        }

        document.add(table);
        document.newPage();

//        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch(); // I know it is ugly, actually you get the actual instance here
//        for (XSSFShape shape : drawing.getShapes()) {
//            if (shape instanceof XSSFPicture picture) {
//                XSSFPictureData xssfPictureData = picture.getPictureData();
//                ClientAnchor anchor = picture.getPreferredSize();
//                if (anchor.getCol1() <= maxCol) {
//                    byte[] data = xssfPictureData.getData();
//                    Image image = Image.getInstance(data);
//                    image.setAbsolutePosition(0,840);
//                    document.add(image);
//                }
//            }
//        }
        //document.newPage();
    }

    private static void setBackgroundColor(Cell cell, PdfPCell cellPdf) {
        // Set background color
        short bgColorIndex = cell.getCellStyle()
                .getFillForegroundColor();
        if (bgColorIndex != IndexedColors.AUTOMATIC.getIndex()) {
            XSSFColor bgColor = (XSSFColor) cell.getCellStyle()
                    .getFillForegroundColorColor();
            if (bgColor != null) {
                byte[] rgb = bgColor.getRGB();
                if (rgb != null && rgb.length == 3) {
                    cellPdf.setBackgroundColor(new RGBColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                } else if (rgb != null) {
                    cellPdf.setBackgroundColor(new RGBColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                }
            }
        }
    }

    private static void setCellAlignment(Cell cell, PdfPCell cellPdf) {
        CellStyle cellStyle = cell.getCellStyle();

        HorizontalAlignment horizontalAlignment = cellStyle.getAlignment();
        VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();

        switch (horizontalAlignment) {
            case LEFT:
                cellPdf.setHorizontalAlignment(Element.ALIGN_LEFT);
                break;
            case CENTER:
                cellPdf.setHorizontalAlignment(Element.ALIGN_CENTER);
                break;
            case JUSTIFY:
            case FILL:
                cellPdf.setVerticalAlignment(Element.ALIGN_JUSTIFIED);
                break;
            case RIGHT:
                cellPdf.setHorizontalAlignment(Element.ALIGN_RIGHT);
                break;
        }

        switch (verticalAlignment) {
            case TOP:
                cellPdf.setVerticalAlignment(Element.ALIGN_TOP);
                break;
            case CENTER:
                cellPdf.setVerticalAlignment(Element.ALIGN_MIDDLE);
                break;
            case JUSTIFY:
                cellPdf.setVerticalAlignment(Element.ALIGN_JUSTIFIED);
                break;
            case BOTTOM:
                cellPdf.setVerticalAlignment(Element.ALIGN_BOTTOM);
                break;
        }
    }


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
                    font.setColor(new RGBColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                }
            }
        }




        short fontSize = cellFont.getFontHeightInPoints();
        font.setSize(fontSize);

        if (cellFont.getBold()) {
            if (cellFont.getItalic()) {
                font.setStyle(Font.BOLDITALIC);
            } else {
                font.setStyle(Font.BOLD);
            }
        } else {
            if (cellFont.getItalic()) {
                font.setStyle(Font.ITALIC);
            }
        }

        if (cellFont.getStrikeout()) {
            font.setStyle(Font.STRIKETHRU);
        }

        if (cellFont.getUnderline() == 1) {
            font.setStyle(Font.UNDERLINE);
        }

        String fontName = cellFont.getFontName();
        if (FontFactory.isRegistered(fontName)) {
            font.setFamily(fontName); // Use extracted font family if supported by iText
        } else {
            //logger.warn("Unsupported font type: {}", fontName);
            // - Use a fallback font (e.g., Helvetica)
            font.setFamily("Calibri");
        }

        return font;
    }

    public static void main(String[] args) throws DocumentException, IOException {
        String excelFilePath = "src/main/resources/excelsample.xlsx";
        String pdfFilePath = "src/main/resources/pdfsample.pdf";
        convertExcelToPDF(excelFilePath, pdfFilePath);
    }
}