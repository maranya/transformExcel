package pe.certicom.com;

import java.io.FileOutputStream;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import pe.certicom.com.base.Configuration;
import pe.certicom.com.base.Constantes;
import pe.certicom.com.excel.Base;

public class Main {

    static Properties config;

    public static void main(String[] args) {
        config = new Configuration().getProperties();
        try {
            Base file = new Base(config.getProperty(Constantes.FILE_IN));
            String rowInit = config.getProperty(Constantes.ROW_INIT);
            String cellInit = config.getProperty(Constantes.CELL_INIT);

            XSSFWorkbook xssf = file.getInstance();
            if (xssf != null) {
                int numDeHojas = xssf.getNumberOfSheets();
                if (numDeHojas > 0) {
                    for (int x = 0; x < numDeHojas; x++) {
                        Constantes.convert(xssf.getSheetAt(x), rowInit, cellInit);
                    }
                } else {
                    throw new Exception();
                }
            }
            FileOutputStream fos = new FileOutputStream(config.getProperty(Constantes.FILE_OUT));
            xssf.write(fos);
            xssf.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

}
