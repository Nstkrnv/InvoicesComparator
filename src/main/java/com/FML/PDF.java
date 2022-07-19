package com.FML;

import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.File;
import java.io.IOException;

public class PDF {
    public static String StringFromPDF(String filename){

        PDFTextStripper pdfStripper = null;
        PDDocument pdDoc = null;
        COSDocument cosDoc = null;
        File file = new File(filename);

        String parsedText = new String();
        try {
            RandomAccessFile randomAccessFile = new RandomAccessFile(file, "r");
            PDFParser parser = new PDFParser(randomAccessFile);

            parser.parse();
            cosDoc = parser.getDocument();
            pdfStripper = new PDFTextStripper();
            pdDoc = new PDDocument(cosDoc);
            pdfStripper.setStartPage(1);
            int quantityOfPages = pdDoc.getNumberOfPages();
            pdfStripper.setEndPage(quantityOfPages);
            parsedText = pdfStripper.getText(pdDoc);
            cosDoc.close();

        } catch (IOException e) {
            System.out.println("Something in PDF went wrong");
            e.printStackTrace();
        }

        return parsedText;
    }
}
