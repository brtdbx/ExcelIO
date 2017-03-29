package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

    public static String myGetCellType(XSSFCell cell) {

        if (cell != null) {
            switch (cell.getCellType()) {
                case XSSFCell.CELL_TYPE_STRING:
                    return "String";
                case XSSFCell.CELL_TYPE_NUMERIC:
                    return "Numeric";
                case XSSFCell.CELL_TYPE_BLANK:
                    return "Blanc";
                case XSSFCell.CELL_TYPE_BOOLEAN:
                    return "Boolean";
                case XSSFCell.CELL_TYPE_ERROR:
                    return "Error";
                case XSSFCell.CELL_TYPE_FORMULA:
                    return "Formula";
                default:
                    return "unknown";
            }

        } else {
            return "cell is null";
        }
    }

    /**
     * Copies the sourceCell's value and styles to the targetCell. Both cells
     * mustn't be null!
     */
    public static void copyCellValueAndStyle(XSSFCell sourceCell, XSSFCell targetCell) {

        XSSFCellStyle newCellStyle;
        newCellStyle = targetCell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
        targetCell.setCellStyle(newCellStyle);

        switch (sourceCell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                targetCell.setCellValue(sourceCell.getRichStringCellValue());
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                targetCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }
    }

    /**
     * Copies the sourceCell's value to the targetCell. Both cells mustn't be
     * null!
     */
    public static void copyCellValue(XSSFCell sourceCell, XSSFCell targetCell) {
        switch (sourceCell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                targetCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }
    }

    /**
     * Get's the sheets's cell identified by rowIndex and colIndex. The sheet
     * mustn't be null. The cell's row and the the cell itself is created, if it
     * doesn't already exist. null!
     */
    public static XSSFCell getCell(XSSFSheet sheet, int rowIndex, int colIndex) {

        if (sheet == null) {
            MyRuntimeException mrtEx = new MyRuntimeException("sheet mustn't be null!");
            Logger.getLogger(ExcelUtils.class.getName()).log(Level.SEVERE, null, mrtEx);
            throw mrtEx;
        }
        XSSFRow row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        XSSFCell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        return cell;
    }

    public static XSSFWorkbook loadReadOnlyWorkbook(String fileName) throws MyRuntimeException {
        File inFile = new File(fileName);
        if (!inFile.exists()) {
            MyRuntimeException mrtEx = new MyRuntimeException(fileName + " existiert nicht", new FileNotFoundException());
            throw mrtEx;
        }
        XSSFWorkbook myWorkBook;
        try {
            myWorkBook = new XSSFWorkbook(inFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelUtils.class.getName()).log(Level.SEVERE, null, ex);
            throw new MyRuntimeException(ex);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ExcelUtils.class.getName()).log(Level.SEVERE, null, ex);
            throw new MyRuntimeException(ex);
        }

        return myWorkBook;
    }

    public static void saveWorkbook(String fileName, XSSFWorkbook workbook) throws IOException, FileNotFoundException {
        FileOutputStream fos = null;
        fos = new FileOutputStream(new File(fileName));
        workbook.write(fos);
        fos.close();
    }

    public static XSSFWorkbook loadReadWriteWorkbook(String fileName) throws FileNotFoundException, IOException {
        FileInputStream fis = new FileInputStream(fileName);
        XSSFWorkbook workbook = null;
        workbook = new XSSFWorkbook(fis);
        fis.close();
        return workbook;
    }

}
