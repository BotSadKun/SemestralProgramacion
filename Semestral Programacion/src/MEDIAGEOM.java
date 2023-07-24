import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class MEDIAGEOM {

    public static void main(String[] args) {
        String nombreArchivo = "C:\\Users\\Juan\\Documents\\programacionsemestral\\FileManager\\excel\\Libro1.xlsx";

        // Calcular media geométrica y guardar el resultado en el archivo Excel
        calcularYGuardarMediaGeometrica(nombreArchivo);
    }

    public static void calcularYGuardarMediaGeometrica(String nombreArchivo) {
        try (FileInputStream fileInputStream = new FileInputStream(nombreArchivo)) {
            // Cargamos el archivo Excel
            Workbook workbook = WorkbookFactory.create(fileInputStream);

            // Si quieres trabajar con una hoja específica, puedes hacerlo así:
            // Sheet hoja = workbook.getSheet("NombreDeLaHoja");
            // Si no, trabajaremos con la primera hoja del libro.
            Sheet hoja = workbook.getSheet("Hoja1");

            // Obtenemos la primera fila donde se encuentran los valores (fila 0)
            Row filaValores = hoja.getRow(0);

            if (filaValores != null) {
                // Variables para realizar el cálculo de la media geométrica
                int cantidadValores = 0;
                double productoValores = 1.0;

                // Recorremos las celdas A1 a G1 para calcular el producto de los valores
                for (Cell celdaValor : filaValores) {
                    if (celdaValor.getCellTypeEnum() == CellType.NUMERIC) {
                        double valor = celdaValor.getNumericCellValue();
                        productoValores *= valor;
                        cantidadValores++;
                    }
                }

                // Calculamos la media geométrica
                double mediaGeometrica = Math.pow(productoValores, 1.0 / cantidadValores);

                // Creamos la fila A2 si no existe
                Row filaResultado = hoja.getRow(1);
                if (filaResultado == null) {
                    filaResultado = hoja.createRow(1);
                }

                // Escribimos el resultado en la celda A2
                Cell celdaResultado = filaResultado.createCell(0);
                celdaResultado.setCellValue(mediaGeometrica);

                // Guardamos los cambios en el archivo
                try (FileOutputStream fileOutputStream = new FileOutputStream(nombreArchivo)) {
                    workbook.write(fileOutputStream);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else {
                System.out.println("La fila de valores no existe.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public static void mediaGeometrica(String nombreArchivo) {
    }
}
