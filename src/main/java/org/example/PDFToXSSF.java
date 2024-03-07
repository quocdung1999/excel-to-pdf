package org.example;

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.pdf.PdfDictionary;
import com.lowagie.text.pdf.PdfReader;
import com.lowagie.text.pdf.PdfTable;
import com.lowagie.text.pdf.parser.PdfTextExtractor;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.Arrays;

public class PDFToXSSF {

    public static void main(String[] args) throws DocumentException, IOException {
        String pdfFilePath = "src/main/resources/excelsample.pdf";
        String excelFilePath = "src/main/resources/pdfsample.xlsx";
        convertPDFToExcel(pdfFilePath, excelFilePath);
    }

    public static void convertPDFToExcel(String pdfFilePath, String excelFilePath) throws IOException, DocumentException {
        ReadPDFDocument(pdfFilePath);
        //Document document = createExcelFile(pdfFilePath);

    }

    public static void ReadPDFDocument(String path) throws IOException {
        PdfReader reader = new PdfReader(path);
        int pages = reader.getNumberOfPages();
        PdfTextExtractor PdfTextExtractor = new PdfTextExtractor(reader);
        //Text
        for (int i = 1; i <= 1; i++) {
            System.out.println("Page: "+ i);

            String a = PdfTextExtractor.getTextFromPage(i);
            System.out.println(a);
        }
    }
}
