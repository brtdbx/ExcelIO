package com.brhacks.excelio;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class App {

    public static void main(String[] args) {

        App controller = new App();
        try {

            controller.loadDataFromExcel();
            // -----------------------
        } catch (Exception ex) {
            Logger.getLogger(App.class.getName()).log(Level.SEVERE, "Schade", ex);
        }
    }

    public void loadDataFromExcel() throws IOException, FileNotFoundException, InvalidFormatException {

        ExcelInteractions1 reader = new ExcelInteractions1();
        String fileName = "c:\\000_br_data\\400_JavaDev\\devProjects\\ExcelIO\\src\\main\\java\\com\\brhacks\\excelio\\decision_tree_example.xlsx";
        XSSFSheet sheet = reader.cloneSheet(fileName);
        reader.dumpSheet(sheet);

    }

    public void readExcel() throws IOException, FileNotFoundException, InvalidFormatException {

        System.out.println("Hallo");
        ExcelInteractions1 reader = new ExcelInteractions1();
        String fileName = "c:\\000_br_data\\400_JavaDev\\devProjects\\ExcelIO\\src\\main\\java\\com\\brhacks\\excelio\\decision_tree_example.xlsx";
        String tmpName = "C:\\000_br_data\\400_JavaDev\\devProjects\\ExcelIO\\src\\main\\java\\com\\brhacks\\excelio\\tmp.xlsx";
        reader.readXLSXFile(fileName);
        reader.writeExcel(tmpName);

        reader.updateExcel(tmpName);
        reader.copyAndMore(fileName, tmpName);
    }
}
