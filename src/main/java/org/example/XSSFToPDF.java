package org.example;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTPicture;

import static org.apache.poi.util.Units.EMU_PER_PIXEL;


public class XSSFToPDF {

    private static final Logger logger = LogManager.getLogger(XSSFToPDF.class);

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

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet worksheet = workbook.getSheetAt(i);
            int rowCount = worksheet.getPhysicalNumberOfRows();
            // Add header with sheet name as title
            Paragraph title = new Paragraph(worksheet.getSheetName(), new Font(Font.FontFamily.HELVETICA, 18, Font.BOLD));
//            System.out.println("Last row number: " + worksheet.getLastRowNum());
//            System.out.println("First row number: " + worksheet.getTopRow());

            XSSFDrawing drawing = (XSSFDrawing) worksheet.createDrawingPatriarch(); // I know it is ugly, actually you get the actual instance here
            for (XSSFShape shape : drawing.getShapes()) {
                if (shape instanceof XSSFPicture picture) {
                    XSSFPictureData xssfPictureData = picture.getPictureData();
                    ClientAnchor anchor = picture.getPreferredSize();
                    int row1 = anchor.getRow1();
                    int row2 = anchor.getRow2();
                    int col1 = anchor.getCol1();
                    int col2 = anchor.getCol2();
                    System.out.println("Row1: " + row1 + " Row2: " + row2);
                    System.out.println("Column1: " + col1 + " Column2: " + col2);
                    System.out.println(anchor.getDx1()/EMU_PER_PIXEL + " " +  anchor.getDx2()/EMU_PER_PIXEL + " " + anchor.getDy1()/EMU_PER_PIXEL + " " + anchor.getDy2()/EMU_PER_PIXEL);
                    // Saving the file
                    String ext = xssfPictureData.suggestFileExtension();
                    byte[] data = xssfPictureData.getData();
                    String filePath = "/Users/o_dung_quoc.p/Work/excel-file-java/image1.png";
                    try (FileOutputStream os = new FileOutputStream(filePath)) {
                        os.write(data);
                        os.flush();
                    }

                }
            }


//            for (int j = 0; j < 10; j++) {
//                System.out.println("Default row height: " + worksheet);
//            }
            title.setSpacingAfter(20f);
            title.setAlignment(Element.ALIGN_CENTER);
            document.add(title);

            createAndAddTable(worksheet, document);
            // Add a new page for each sheet (except the last one)
            if (i < workbook.getNumberOfSheets() - 1) {
                document.newPage();
            }
        }

        document.close();
        workbook.close();
    }

    private static void createAndAddTable(Sheet worksheet, Document document) throws DocumentException, IOException {
        PdfPTable table = new PdfPTable(worksheet.getRow(0)
                .getPhysicalNumberOfCells());
        table.setWidthPercentage(100);
        addTableData(worksheet, table);
        document.add(table);
    }


    public static String getCellText(Cell cell) {
        String cellValue;
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case NUMERIC:
                cellValue = String.valueOf(BigDecimal.valueOf(cell.getNumericCellValue()));
                break;
            case BLANK:
            default:
                cellValue = "";
                break;
        }
        return cellValue;
    }

    private static void addTableData(Sheet worksheet, PdfPTable table) throws DocumentException, IOException {
        for (Row row : worksheet) {
            int currIndex = 0;
            for (int i = 0; i < row.getPhysicalNumberOfCells(); currIndex++) {
                Cell cell = row.getCell(currIndex);
                //System.out.println(row.getRowNum()+ ","+ currIndex);
                PdfPCell cellPdf = null;
                if (cell != null) {
                    String cellValue = getCellText(cell);
                    cellPdf = new PdfPCell(new Phrase(cellValue, util.getCellStyle(cell)));
                    setBackgroundColor(cell, cellPdf);
                    setCellAlignment(cell, cellPdf);
                    i++;
                } else {
                    cellPdf = new PdfPCell(new Phrase("", null));
                }
                table.addCell(cellPdf);
            }
        }
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
                    cellPdf.setBackgroundColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
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



    public static void main(String[] args) throws DocumentException, IOException {
        String excelFilePath = "src/main/resources/Book1.xlsx";
        String pdfFilePath = "src/main/resources/pdfsample.pdf";
        convertExcelToPDF(excelFilePath, pdfFilePath);
    }
}