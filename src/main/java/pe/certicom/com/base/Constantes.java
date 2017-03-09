package pe.certicom.com.base;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Constantes {

    public static final String FILE = "D:\\test\\convert_excel.properties";

    public static final String FILE_IN = "name.fileIn";
    public static final String FILE_OUT = "name.fileOut";
    public static final String ROW_INIT = "num.rowInit";
    public static final String CELL_INIT = "num.cellInit";

    public static void convert(XSSFSheet sheet, String rowInit, String cellInit) {
        String[] rows = null, cells = null;
        if (rowInit != null && !rowInit.isEmpty()) {
            rows = rowInit.split(",");
        }
        if (cellInit != null && !cellInit.isEmpty()) {
            cells = cellInit.split(",");
        }
        if (rows == null || cells == null) {
            System.out.println("No existe configuracion para filas o columnas de inicio");
        }

        Iterator<Row> rowIterator = sheet.iterator();
        CellStyle celNumericStyle = null;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (existeRow(rows, String.valueOf(row.getRowNum()))) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell celda = cellIterator.next();
                    if (existeCell(cells, String.valueOf(celda.getColumnIndex()))) {
                        switch (celda.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                String value = celda.getStringCellValue().trim();

                                System.out.println(celda.getStringCellValue());
                                if (value.length() == 1 && value.equals(".")) {
                                    value = "";
                                    celda.setCellValue(value);
                                } else {
                                    String newValue = addValueReplace(value, "-", "-");
                                    if (celNumericStyle != null) {
                                        celNumericStyle.setBorderBottom(CellStyle.BORDER_NONE);
                                        celNumericStyle.setBorderLeft(CellStyle.BORDER_NONE);
                                        celNumericStyle.setBorderRight(CellStyle.BORDER_NONE);
                                        celNumericStyle.setBorderTop(CellStyle.BORDER_NONE);

                                        celda.setCellStyle(celNumericStyle);
                                    }
                                    celda.setCellValue(Double.parseDouble(newValue));
                                }
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                System.out.println(celda.getNumericCellValue());
                                if (celNumericStyle == null) {
                                    celNumericStyle = celda.getCellStyle();
                                }
                                break;
                        }
                    }
                }
            }
        }
    }

    protected static Boolean existeCell(String[] values, String search) {
        if (values.length == 1 && Integer.parseInt(search) >= Integer.parseInt(values[0])) {
            return Boolean.TRUE;
        } else {
            for (int x = 0; x < values.length; x++) {
                if (values[x].equals(search)) {
                    return Boolean.TRUE;
                }
            }
        }
        return Boolean.FALSE;
    }

    protected static Boolean existeRow(String[] values, String search) {
        if (values.length == 1 && Integer.parseInt(values[0]) <= Integer.parseInt(search)) {
            return Boolean.TRUE;
        } else {
            for (int x = 0; x < values.length; x++) {
                if (values[x].equals(search)) {
                    return Boolean.TRUE;
                }
            }
        }
        return Boolean.FALSE;
    }

    protected static String replace(String old, String value) {
        String str = "";
        if (value != null && !value.isEmpty()) {
            for (int x = 0; x < old.length(); x++) {
                String valorActual = old.substring(x, x + 1);
                if (!valorActual.equals(",") && !valorActual.equals(value)) {
                    str += valorActual;
                }
            }
        } else {
            str = old;
        }

        return str;
    }

    protected static String addValueReplace(String old, String value, String add) {
        boolean first = false;
        if (old.contains(value)) {
            first = true;
        }
        String str = replace(old, value);
        if (first && add != null && !add.isEmpty()) {
            str = add + str;
        }
        return str;
    }
}
