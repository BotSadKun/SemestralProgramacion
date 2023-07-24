import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class MediaGeometricaExcel {

    public static void main(String[] args) {
        String nombreArchivo = "C:\\Users\\Juan\\Documents\\programacionsemestral\\FileManager\\excel\\Libro1.xlsx";
        calcularYGuardarMediaGeometrica(nombreArchivo);
    }

    public static void calcularYGuardarMediaGeometrica(String nombreArchivo) {
        try (FileInputStream fileInputStream = new FileInputStream(nombreArchivo)) {
            Workbook workbook = WorkbookFactory.create(fileInputStream);

            Sheet hoja = workbook.getSheet("Hoja1"); //Importante

            Cell celdaA1 = hoja.getRow(0).getCell(0); //Importante
            Cell celdaB1 = hoja.getRow(0).getCell(1);
            Cell celdaC1 = hoja.getRow(0).getCell(2);
            Cell celdaD1 = hoja.getRow(0).getCell(3);
            Cell celdaE1 = hoja.getRow(0).getCell(4);
            Cell celdaF1 = hoja.getRow(0).getCell(5);
            Cell celdaG1 = hoja.getRow(0).getCell(6);

            if (celdaA1.getCellTypeEnum() == CellType.NUMERIC && //Importante
                celdaB1.getCellTypeEnum() == CellType.NUMERIC && 
                celdaC1.getCellTypeEnum() == CellType.NUMERIC &&
                celdaD1.getCellTypeEnum() == CellType.NUMERIC &&
                celdaE1.getCellTypeEnum() == CellType.NUMERIC &&
                celdaF1.getCellTypeEnum() == CellType.NUMERIC &&
                celdaG1.getCellTypeEnum() == CellType.NUMERIC) {

                double productoValores = 
                                        celdaA1.getNumericCellValue() * //Importante
                                        celdaB1.getNumericCellValue() *
                                        celdaC1.getNumericCellValue() *
                                        celdaD1.getNumericCellValue() *
                                        celdaE1.getNumericCellValue() *
                                        celdaF1.getNumericCellValue() *
                                        celdaG1.getNumericCellValue()
                ;

                double mediaGeometrica = Math.pow(productoValores, 1.0 / 7); //Importante

                Row filaResultado = hoja.getRow(1);
                if (filaResultado == null) {
                    filaResultado = hoja.createRow(1);
                }

                Cell celdaResultado = filaResultado.createCell(0);
                celdaResultado.setCellValue(mediaGeometrica);

                try (FileOutputStream fileOutputStream = new FileOutputStream(nombreArchivo)) {
                    workbook.write(fileOutputStream);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else {
                System.out.println("Las celdas deben contener valores numéricos.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            e.printStackTrace();
        }
    }
}
