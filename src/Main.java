import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.regex.*;

public class Main {
    public static void main(String[] args) throws Exception {
        String os = System.getProperty("os.name").toLowerCase();
        Path carpeta;
        if (os.contains("win")) {
            carpeta = Paths.get("Consultores");
        } else {
            carpeta = Paths.get("Consultores");
        }
        // Recorrer todos los archivos XLSX en la carpeta
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(carpeta, "*.xlsx")) {
            int i = 0;
            for (Path archivo : stream) {
                //System.out.println("ciclo");
                i++;
                try (FileInputStream fis = new FileInputStream(archivo.toFile());
                     Workbook workbook = new XSSFWorkbook(fis)) {

                    Sheet hoja = workbook.getSheet("JUNIO 2025"); // o usar workbook.getSheetAt(0)
                    System.out.println(archivo.toString());

                    String partidaActual = null;
                    Pattern patronPartida = Pattern.compile("^\\d+.*"); // empieza con número

                    for (Row fila : hoja) {
                        if (fila.getRowNum() < 10) continue; // Saltar encabezado

                        Cell celda0 = fila.getCell(0); // la celda donde estaría la partida
                        if (celda0 != null && celda0.getCellTypeEnum() == CellType.STRING) {
                            String texto = celda0.getStringCellValue().trim();

                            // Comprobar si parece una partida
                            Matcher matcher = patronPartida.matcher(texto);
                            if (matcher.find()) {
                                // Verificar si tiene fondo gris
                                CellStyle style = celda0.getCellStyle();
                                if (style.getFillForegroundColorColor() != null) {
                                    partidaActual = texto;
                                    System.out.println("PARTIDA DETECTADA: " + partidaActual);
                                    continue; // no es una fila de persona
                                }
                            }
                        }

                        // Si es una fila de persona (no partida)
                        Cell nombre = fila.getCell(6); // columna G
                        Cell cargo = fila.getCell(7);  // columna H
                        Cell monto = fila.getCell(11); // columna L

                        if (nombre != null && nombre.getCellTypeEnum() == CellType.STRING && !nombre.getStringCellValue().isEmpty()) {
                            System.out.println("Partida: " + partidaActual);
                            System.out.println("Nombre: " + nombre.getStringCellValue());
                            System.out.println("Cargo: " + (cargo != null ? cargo.getStringCellValue() : "Sin cargo"));
                            System.out.println("Monto: " + (monto != null ? monto.getNumericCellValue() : 0));
                            System.out.println("---");
                        }
                    }
                }
            }
            System.out.println(i + " Procesados");
        }
    }
}