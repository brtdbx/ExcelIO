package com.brhacks.excelio;

import utils.MyRuntimeException;
import utils.ExcelUtils;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelInteractions1 {

    public void dumpSheet(XSSFSheet sheet) {

        if (sheet != null) {

            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                XSSFRow srcRow = sheet.getRow(i);

                for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {

                    if (j < 0) {
                        break;
                    }
//                    System.out.println(srcRow.getFirstCellNum());
//                    System.out.println(srcRow.getLastCellNum());


                    XSSFCell cell = srcRow.getCell(j);
                    if (cell != null) {
                        System.out.println(""  + i + " " + j + " " + cell.toString() + " Type:" + ExcelUtils.myGetCellType(cell));
                    }
                }
            }
            System.out.println("Done");
        } else {
            System.out.println("sheet is empty");
        }

    }

    public XSSFSheet cloneSheet(String inputFilename) throws IOException {

        // File inFile = new File(inputFilename);
        XSSFSheet outputSheet;
        try (XSSFWorkbook inputWorkbook = ExcelUtils.loadReadOnlyWorkbook(inputFilename)) {
            XSSFSheet inputSheet = inputWorkbook.getSheetAt(0);
            XSSFWorkbook outputWorkbook = new XSSFWorkbook();
            outputSheet = outputWorkbook.createSheet();
            for (int i = inputSheet.getFirstRowNum(); i <= inputSheet.getLastRowNum(); i++) {
                XSSFRow srcRow = inputSheet.getRow(i);
                XSSFRow destRow = outputSheet.createRow(i);

                for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {

                    if (j < 0) {
                        break;
                    }
//                    System.out.println(srcRow.getFirstCellNum());
//                    System.out.println(srcRow.getLastCellNum());

//                    System.out.println("i,j " + i + " " + j);

                    XSSFCell oldCell = srcRow.getCell(j);
                    XSSFCell newCell = destRow.getCell(j);
                    if (oldCell != null) {
                        if (newCell == null) {
                            newCell = destRow.createCell(j);
                        }
                        ExcelUtils.copyCellValueAndStyle(oldCell, newCell);
                    }
                }
            }
        }

        System.out.println("Done");
        return outputSheet;

    }

    public void copyAndMore(String inputFilename, String outputFilename) throws MyRuntimeException {

        try {

            // File inFile = new File(inputFilename);
            XSSFWorkbook inputWorkbook = ExcelUtils.loadReadOnlyWorkbook(inputFilename);
            XSSFSheet inputSheet = inputWorkbook.getSheetAt(0);

            XSSFWorkbook outputWorkbook = new XSSFWorkbook();
            XSSFSheet outputSheet = outputWorkbook.createSheet("meinSheet2");
            for (int i = inputSheet.getFirstRowNum(); i <= inputSheet.getLastRowNum(); i++) {

                XSSFRow srcRow = inputSheet.getRow(i);
                XSSFRow destRow = outputSheet.createRow(i);

                for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {

                    if (j < 0) {
                        break;
                    }
                    System.out.println(srcRow.getFirstCellNum());
                    System.out.println(srcRow.getLastCellNum());

                    System.out.println("i,j " + i + " " + j);

                    XSSFCell oldCell = srcRow.getCell(j);
                    XSSFCell newCell = destRow.getCell(j);
                    if (oldCell != null) {
                        if (newCell == null) {
                            newCell = destRow.createCell(j);
                        }
                        ExcelUtils.copyCellValueAndStyle(oldCell, newCell);
                    }
                }
            }
            setFooter(outputSheet);

            ExcelUtils.saveWorkbook(outputFilename, outputWorkbook);

            System.out.println("Done");

        } catch (IOException ex) {
            throw new MyRuntimeException("Abbruch:", ex);
        }

    }

    public void readXLSXFile(String fileName) {

        XSSFWorkbook myWorkBook;
        myWorkBook = ExcelUtils.loadReadOnlyWorkbook(fileName);

        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        System.out.println(mySheet.getFirstRowNum() + " " + mySheet.getLastRowNum());

// Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();
// Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // For each row, iterate through each columns
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {

                Cell cell = cellIterator.next();

                // cell.getColumnIndex()
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "\t");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t");
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue() + "\t");
                        break;
                    default:

                }
            }
            System.out.println("");

        }
        try {
            myWorkBook.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelInteractions1.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void updateExcel(String fileName) {
        try {

            XSSFWorkbook workbook = ExcelUtils.loadReadWriteWorkbook(fileName);

//Create First Row
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow row1 = sheet.getRow(1);
            XSSFCell cell1 = row1.getCell(1);
            cell1.setCellValue("Updated!");

            ExcelUtils.saveWorkbook(fileName, workbook);

            System.out.println("Done");
        } catch (IOException ex) {
            throw new MyRuntimeException("Abbruch:", ex);
        }
    }

    public void writeExcel(String fileName) {
        FileOutputStream fos = null;
        try {
//            String fileName = "c:\\400_JavaDev\\devProjects\\ExcelIO\\src\\main\\java\\com\\brhacks\\excelio\\tmp.xlsx";
            //FileInputStream fis = new FileInputStream(new File(fileName));
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("meinSheet");
            //Create First Row
            XSSFRow row1 = sheet.createRow(0);
            XSSFCell r1c1 = row1.createCell(0);
            r1c1.setCellValue("Emd Id");
            XSSFCell r1c2 = row1.createCell(1);
            r1c2.setCellValue("NAME");
            XSSFCell r1c3 = row1.createCell(2);
            r1c3.setCellValue("AGE");
            //Create Second Row
            XSSFRow row2 = sheet.createRow(1);
            XSSFCell r2c1 = row2.createCell(0);
            r2c1.setCellValue("1");
            XSSFCell r2c2 = row2.createCell(1);
            r2c2.setCellValue("Ram");
            XSSFCell r2c3 = row2.createCell(2);
            r2c3.setCellValue("20");
            //Create Third Row
            XSSFRow row3 = sheet.createRow(2);
            XSSFCell r3c1 = row3.createCell(0);
            r3c1.setCellValue("2");
            XSSFCell r3c2 = row3.createCell(1);
            r3c2.setCellValue("Shyam");
            XSSFCell r3c3 = row3.createCell(2);
            r3c3.setCellValue("25");
            //fis.close();
            fos = new FileOutputStream(new File(fileName));
            workbook.write(fos);
            fos.close();
            System.out.println("Done");
        } catch (IOException ex) {
            throw new MyRuntimeException("Abbruch:", ex);
        } finally {
            try {
                fos.close();
            } catch (IOException ex) {
                throw new MyRuntimeException("Abbruch:", ex);
            }
        }
    }

    private void setFooter(Sheet sheet) {
        Footer footer = sheet.getFooter();
        footer.setRight("Seite " + HeaderFooter.page() + " / " + HeaderFooter.numPages());
    }

}
