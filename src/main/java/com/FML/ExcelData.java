package com.FML;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.*;
//import org.apache.poi.ss.usermodel.*;

import java.io.*;

import java.util.Iterator;



import static java.lang.Integer.valueOf;


public class ExcelData {


    public static void parsingExcel(String filename, String bigString){

        bigString = bigString.replaceAll("\r\n", " ");

        try {

            InputStream fs = new FileInputStream(filename);
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            fs.close();
            File file = new File(filename);
            FileOutputStream fileOut = new FileOutputStream(file.getParentFile()+"\\new.xlsx");


            XSSFSheet sheet = wb.getSheet(" week");

            XSSFSheet sheetComparison= wb.createSheet("Comparison");
            XSSFRow rowHead = sheetComparison.createRow(0);

            XSSFCell cell0 = rowHead.createCell(0);
            cell0.setCellValue(new XSSFRichTextString("Invoice"));

            XSSFCell cell1 = rowHead.createCell(1);
            cell1.setCellValue(new XSSFRichTextString("Date of issue"));

            XSSFCell cell2 = rowHead.createCell(2);
            cell2.setCellValue(new XSSFRichTextString("Number of places"));

            XSSFCell cell3 = rowHead.createCell(3);
            cell3.setCellValue(new XSSFRichTextString("Quantita Imballaggi"));

            XSSFCell cell4 = rowHead.createCell(4);
            cell4.setCellValue(new XSSFRichTextString("Weight Brutto"));

            XSSFCell cell5 = rowHead.createCell(5);
            cell5.setCellValue(new XSSFRichTextString("P Lordo"));

            XSSFCell cell6 = rowHead.createCell(6);
            cell6.setCellValue(new XSSFRichTextString("Weight Net"));

            XSSFCell cell7 = rowHead.createCell(7);
            cell7.setCellValue(new XSSFRichTextString("P Netto"));

            XSSFCell cell8 = rowHead.createCell(8);
            cell8.setCellValue(new XSSFRichTextString("Volume Excel"));

            XSSFCell cell9 = rowHead.createCell(9);
            cell9.setCellValue(new XSSFRichTextString("Volume PDF"));

            XSSFCell cell10 = rowHead.createCell(10);
            cell10.setCellValue(new XSSFRichTextString("Number of items"));

            XSSFCell cell11 = rowHead.createCell(11);
            cell11.setCellValue(new XSSFRichTextString("Pezzi"));

            XSSFCell cell12 = rowHead.createCell(12);
            cell12.setCellValue(new XSSFRichTextString("Value (Price)"));

            XSSFCell cell13 = rowHead.createCell(13);
            cell13.setCellValue(new XSSFRichTextString("Valore (Price)"));

            XSSFCell cell14 = rowHead.createCell(14);
            cell14.setCellValue(new XSSFRichTextString("Content"));

            XSSFCell cell15 = rowHead.createCell(15);
            cell15.setCellValue(new XSSFRichTextString("Shop"));

            XSSFCell cell16 = rowHead.createCell(16);
            cell16.setCellValue(new XSSFRichTextString("Comments"));

            wb.write(fileOut);
            fileOut.close();
            //  fs.close();


            Iterator rowIter = sheet.rowIterator();


            int currentIndexOfBigString = 0;
            for (int j = 0; rowIter.hasNext(); j++) {

             /*   if (rowHead.getCell(1) == null){
                    fileOut.close();
                    fs.close();
                    break;
                }*/
                if(!rowIter.hasNext()){
                    break;
                }
                XSSFRow row = (XSSFRow) rowIter.next();
                if (row.getRowNum()==0 || row.getRowNum()==1){
                    continue;
                }
               /* XSSFCell id = row.getCell(0);
                if (id==null){
                    break;
                }*/
                XSSFCell invoice = row.getCell(1);
                XSSFCell dateOfIssue = row.getCell(2);
                XSSFCell NumberOfPlaces = row.getCell(3);
                XSSFCell weightBrut = row.getCell(4);
                XSSFCell weightNet = row.getCell(5);
                XSSFCell volumeExcel = row.getCell(6);
                XSSFCell numberOfItems = row.getCell(7);
                XSSFCell priceExcel = row.getCell(8);
                XSSFCell content = row.getCell(9);
                XSSFCell shop = row.getCell(10);
                XSSFCell comments = row.getCell(11);

                String quantita;
                String pLordo;
                String pNetto;
                String volumePDF;
                int pezzi=0;
                String pricePDF;

                if(invoice==null){
                    // fileOut.close();
                    // fs.close();
                    break;
                }

                StringBuilder invoiceSB = new StringBuilder(invoice.toString());
                invoiceSB.deleteCharAt(4);
                invoiceSB.deleteCharAt(5);

                FileOutputStream fileOut0 = new FileOutputStream(file.getParentFile()+"\\new.xlsx");

                XSSFRow rowComp = sheetComparison.createRow((short)j);

                currentIndexOfBigString = bigString.lastIndexOf(String.valueOf(invoiceSB));
                if (currentIndexOfBigString==-1){
                    XSSFCell cellNon = rowComp.createCell(0);
                    cellNon.setCellValue(String.valueOf(invoice));


                    XSSFCell cellNon2 = rowComp.createCell(17);
                    cellNon2.setCellValue("отсутствует в PDF файле");
                    System.out.println("строка отсутствует в PDF файле");

                    wb.write(fileOut0);
                    fileOut0.close();

                    continue;
                }


                wb.write(fileOut0);
                fileOut0.close();


                currentIndexOfBigString = bigString.indexOf("/", currentIndexOfBigString);
                int endIndex0 = currentIndexOfBigString;

                for (int i = currentIndexOfBigString; ;i++){
                    if (bigString.charAt(i) == ' '){
                        endIndex0 = i;
                    }
                    if (Character.isLetter(bigString.charAt(i+2))){
                        quantita = bigString.substring(endIndex0+1, i+1); //    10696840/1 1 Ra       88/1 0 10 C
                        currentIndexOfBigString = i+3;
                        break;
                    }
                }


                for (int i = currentIndexOfBigString; ; i++){
                    if (Character.isDigit(bigString.charAt(i))){
                        currentIndexOfBigString=i;
                        break;
                    }
                }
                int indexSpaceBeforeLordo=0;
                for (int i = currentIndexOfBigString; ;){
                    int endIndex = bigString.indexOf(' ', i);

                    String pStr = bigString.substring(i, endIndex);
                    if (pStr.contains(".")){
                        break;
                    }

                    if(pStr.length()==1){i+=2;}
                    if(pStr.length()==2){i+=3;}
                    if(pStr.length()==3){i+=4;}

                    int p = valueOf(pStr);
                    pezzi+=p;
                    indexSpaceBeforeLordo = endIndex;
                }
                currentIndexOfBigString = ++indexSpaceBeforeLordo;

                int endIndex = bigString.indexOf(' ', currentIndexOfBigString);
                pLordo = bigString.substring(currentIndexOfBigString, endIndex);
                currentIndexOfBigString+=pLordo.length()+1;

                endIndex = bigString.indexOf(' ', currentIndexOfBigString);
                pNetto = bigString.substring(currentIndexOfBigString, endIndex);
                currentIndexOfBigString+=pNetto.length()+1;

                endIndex = bigString.indexOf(' ', currentIndexOfBigString);
                pricePDF = bigString.substring(currentIndexOfBigString, endIndex);
                currentIndexOfBigString+=pricePDF.length()+1;

                currentIndexOfBigString+=4;
                endIndex = bigString.indexOf(' ', currentIndexOfBigString);
                volumePDF = bigString.substring(currentIndexOfBigString, endIndex);


                FileOutputStream fileOut1 = new FileOutputStream(file.getParentFile()+"\\new.xlsx");
                //XSSFRow rowComp = sheetComparison.createRow((short)j);


//                java.awt.Color COLOR_light_gray  = new java.awt.Color(252, 3, 73, 255);//252/3/73
                XSSFCellStyle style = wb.createCellStyle();
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//                style.setFillForegroundColor(new XSSFColor(COLOR_light_gray));

                style.setFillForegroundColor(new XSSFColor(new java.awt.Color(252, 3, 73, 255), null));


                try{
                    XSSFCell cell00 = rowComp.createCell(0);
                    cell00.setCellValue(String.valueOf(invoice));
                    sheetComparison.autoSizeColumn(0);

                    XSSFCell cell11a = rowComp.createCell(1);
                    cell11a.setCellValue(String.valueOf(dateOfIssue));
                    sheetComparison.autoSizeColumn(1);

                    XSSFCell cell22 = rowComp.createCell(2);
                    double NumberOfPlacesDouble = Double.parseDouble(String.valueOf(NumberOfPlaces));
                    cell22.setCellValue(NumberOfPlacesDouble);
                    sheetComparison.autoSizeColumn(2);

                    XSSFCell cell33 = rowComp.createCell(3);
                    double quantitaDouble = Double.parseDouble(String.valueOf(quantita));
                    cell33.setCellValue(quantitaDouble);
                    sheetComparison.autoSizeColumn(3);

                    XSSFCell cell44 = rowComp.createCell(4);
                    double weightBrutDouble = Double.parseDouble(String.valueOf(weightBrut));
                    cell44.setCellValue(weightBrutDouble);
                    sheetComparison.autoSizeColumn(4);

                    XSSFCell cell55 = rowComp.createCell(5);
                    double pLordoDouble = Double.parseDouble(pLordo);
                    cell55.setCellValue(pLordoDouble);
                    sheetComparison.autoSizeColumn(5);

                    XSSFCell cell66 = rowComp.createCell(6);
                    double weightNetDouble = Double.parseDouble(String.valueOf(weightNet));
                    cell66.setCellValue(weightNetDouble);
                    sheetComparison.autoSizeColumn(6);

                    XSSFCell cell77 = rowComp.createCell(7);
                    double pNettoDouble = Double.parseDouble(pNetto);
                    cell77.setCellValue(pNettoDouble);
                    sheetComparison.autoSizeColumn(7);

                    XSSFCell cell88 = rowComp.createCell(8);
                    double DoublevolumeExcel = Double.parseDouble(String.valueOf(volumeExcel));
                    cell88.setCellValue(DoublevolumeExcel);
                    sheetComparison.autoSizeColumn(8);

                    XSSFCell cell99 = rowComp.createCell(9);
                    double volumePDFDouble = Double.parseDouble(volumePDF);
                    cell99.setCellValue(volumePDFDouble);
                    sheetComparison.autoSizeColumn(9);

                    XSSFCell cell1010 = rowComp.createCell(10);
                    cell1010.setCellValue(Double.parseDouble(String.valueOf(numberOfItems)));
                    sheetComparison.autoSizeColumn(10);

                    XSSFCell cell1111 = rowComp.createCell(11);
                    cell1111.setCellValue(Double.parseDouble(String.valueOf(pezzi)));
                    sheetComparison.autoSizeColumn(11);

                    XSSFCell cell1212 = rowComp.createCell(12);
                    double DoublepriceExcel = Double.parseDouble(String.valueOf(priceExcel));
                    cell1212.setCellValue(DoublepriceExcel);
                    sheetComparison.autoSizeColumn(12);

                    XSSFCell cell1313 = rowComp.createCell(13);
                    double pricePDFDouble = Double.parseDouble(pricePDF);
                    cell1313.setCellValue(pricePDFDouble);
                    sheetComparison.autoSizeColumn(13);

                    XSSFCell cell1414 = rowComp.createCell(14);
                    cell1414.setCellValue(String.valueOf(content));
                    sheetComparison.autoSizeColumn(14);

                    XSSFCell cell1515 = rowComp.createCell(15);
                    cell1515.setCellValue(String.valueOf(shop));
                    sheetComparison.autoSizeColumn(15);

                    XSSFCell cell1616 = rowComp.createCell(16);
                    cell1616.setCellValue(String.valueOf(comments));
                    sheetComparison.autoSizeColumn(16);

                    if (quantitaDouble!=NumberOfPlacesDouble){
                        cell22.setCellStyle(style);
                        cell33.setCellStyle(style);
                        XSSFCell cell1717 = rowComp.createCell(17);
                        cell1717.setCellValue("не соответствует");
                        sheetComparison.autoSizeColumn(17);
                    }

                    if (Math.abs(weightBrutDouble)-Math.abs(pLordoDouble) >= 0.1){
                        cell44.setCellStyle(style);
                        cell55.setCellStyle(style);
                        XSSFCell cell1717 = rowComp.createCell(17);
                        cell1717.setCellValue("не соответствует");
                        sheetComparison.autoSizeColumn(17);
                    }
                    if (Math.abs(weightNetDouble)-Math.abs(pNettoDouble) >= 0.1){
                        cell66.setCellStyle(style);
                        cell77.setCellStyle(style);
                        XSSFCell cell1717 = rowComp.createCell(17);
                        cell1717.setCellValue("не соответствует");
                        sheetComparison.autoSizeColumn(17);
                    }
                    if (Math.abs(DoublevolumeExcel)-Math.abs(volumePDFDouble) >= 0.001){
                        cell88.setCellStyle(style);
                        cell99.setCellStyle(style);
                        XSSFCell cell1717 = rowComp.createCell(17);
                        cell1717.setCellValue("не соответствует");
                        sheetComparison.autoSizeColumn(17);
                    }
                    System.out.println("строки идентичны");
                    if(Math.abs(DoublepriceExcel)-Math.abs(pricePDFDouble) >= 0.01){
                        cell1212.setCellStyle(style);
                        cell1313.setCellStyle(style);
                        XSSFCell cell1717 = rowComp.createCell(17);
                        cell1717.setCellValue("не соответствует");
                        sheetComparison.autoSizeColumn(17);
                    }
                }catch (Exception e){
                    System.out.println("не могу пропарсить");
                }

                wb.write(fileOut1);
                fileOut1.close();
            }

        }catch (Exception e) {
            System.out.println("Something in EXcel went wrong");
        }
    }
}
