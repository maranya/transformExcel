package pe.certicom.com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Base {

    FileInputStream is;

    XSSFWorkbook workbook;

    public Base(String srcFile) {
        try {
            this.is = new FileInputStream(new File(srcFile));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    public XSSFWorkbook getInstance() throws Exception {
        if (this.is == null) {
            throw new Exception("Error al leer archivo");
        } else {
            workbook = new XSSFWorkbook(this.is);
        }
        return workbook;
    }
}
