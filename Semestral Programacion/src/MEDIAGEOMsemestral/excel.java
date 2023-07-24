package MEDIAGEOMsemestral;
public class excel {
    public static void main(String[] args) {
        String nombreArchivo = "C:\\Users\\Juan\\Documents\\programacionsemestral\\Semestral Programacion\\excel\\Libro1.xlsx";
        
        mediaGeom calcularMEDIAGEOM = new mediaGeom();
        calcularMEDIAGEOM.calculoExcelMediaGeom(nombreArchivo);
    }
}
