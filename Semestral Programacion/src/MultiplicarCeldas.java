import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.apache.poi.ss.usermodel.*;


public class MultiplicarCeldas {

    public static void main(String[] args) {
        String nombreArchivo = "C:\\Users\\Juan\\Documents\\programacionsemestral\\FileManager\\excel\\Libro1.xlsx";

        // Cerrar Excel si está abierto
        cerrarExcel();

        // Multiplicar celdas y guardar el resultado
        multiplicarCeldas(nombreArchivo);
    }

    public static void multiplicarCeldas(String nombreArchivo) {
        try (FileInputStream fileInputStream = new FileInputStream(nombreArchivo)) {
            // Cargamos el archivo Excel
            Workbook workbook = WorkbookFactory.create(fileInputStream);

            // Si quieres trabajar con una hoja específica, puedes hacerlo así:
            // Sheet hoja = workbook.getSheet("NombreDeLaHoja");
            // Si no, trabajaremos con la primera hoja del libro.
            Sheet hoja = workbook.getSheetAt(0);

            // Obtenemos la primera fila (índice 0)
            Row filaA1 = hoja.getRow(0);

            if (filaA1 != null) {
                // Obtenemos la celda A1
                Cell celdaA1 = filaA1.getCell(0);

                // Verificamos si la celda contiene un valor numérico
                if (celdaA1.getCellTypeEnum() == CellType.NUMERIC) {
                    double valorA1 = celdaA1.getNumericCellValue();

                    // Realizamos la multiplicación
                    double resultado = valorA1 * 2; // Cambia 2 por el número por el que deseas multiplicar

                    // Escribimos el resultado en la celda B1
                    Cell celdaB1 = filaA1.createCell(1); // Creamos la celda B1 si no existe
                    celdaB1.setCellValue(resultado);

                    // Guardamos los cambios en el archivo
                    try (FileOutputStream fileOutputStream = new FileOutputStream(nombreArchivo)) {
                        workbook.write(fileOutputStream);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                } else {
                    System.out.println("La celda A1 no contiene un valor numérico.");
                }
            } else {
                System.out.println("La fila 1 no existe.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public static void cerrarExcel() {
        try {
            // Ruta del archivo VBS que cierra Excel
            String rutaVBS = "C:\\Users\\Juan\\Documents\\programacionsemestral\\FileManager\\excel\\CerrarExcel.vbs";

            // Verificar si el archivo VBS existe
            if (Files.exists(Paths.get(rutaVBS))) {
                // Ejecutar el script VBS para cerrar Excel
                ProcessBuilder processBuilder = new ProcessBuilder("wscript.exe", rutaVBS);
                processBuilder.start().waitFor();
            } else {
                System.out.println("El archivo VBS para cerrar Excel no se encontró.");
            }
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}
