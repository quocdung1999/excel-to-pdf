package org.example;

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.pdf.PdfReader;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;

public class PDFToXSSF {

    public static void main(String[] args) throws DocumentException, IOException {
        String pdfFilePath = "src/main/resources/Book1.pdf";
        String excelFilePath = "src/main/resources/excelsample.xlsx";
        convertPDFToExcel(excelFilePath, pdfFilePath);
    }

    public static void convertPDFToExcel(String pdfFilePath, String excelFilePath) throws IOException, DocumentException {
        //Workbook workbook = ReadPDFDocument(excelFilePath);
        //Document document = createExcelFile(pdfFilePath);

    }

    public static void ReadPDFDocument(String path) throws IOException {
        PdfReader reader = new PdfReader("sample.pdf");
        int pages = reader.getNumberOfPages();
    }
}
