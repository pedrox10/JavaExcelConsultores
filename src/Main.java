import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.regex.*;

public class Main {
    public static void main(String[] args) throws Exception {
        Path carpeta = Paths.get("Consultores");
        int i = 0;
        int index_funcionarios = 0;
        int index_partida = 0;

        // Abrimos el archivo de salida solo una vez
        FileInputStream fisOut = new FileInputStream("Resultados/resultados.xlsx");
        Workbook workbookOut = new XSSFWorkbook(fisOut);
        fisOut.close(); // cerrar después de cargar

        // Preparamos hojas y estilos
        Sheet hojaPartidas = workbookOut.getSheetAt(3);
        CellStyle estiloConBordes = workbookOut.createCellStyle();
        estiloConBordes.setBorderTop(BorderStyle.THIN);
        estiloConBordes.setBorderBottom(BorderStyle.THIN);
        estiloConBordes.setBorderLeft(BorderStyle.THIN);
        estiloConBordes.setBorderRight(BorderStyle.THIN);

        Pattern patronPartida = Pattern.compile("^\\d+.*"); // empieza con número

        try (DirectoryStream<Path> stream = Files.newDirectoryStream(carpeta, "*.xlsx")) {
            for (Path archivo : stream) {
                i++;
                try (FileInputStream fis = new FileInputStream(archivo.toFile());
                     Workbook workbook = new XSSFWorkbook(fis)) {

                    Sheet hoja = workbook.getSheet("JUNIO 2025");
                    System.out.println("Procesando: " + archivo.getFileName());

                    String partidaActual = null;

                    for (Row fila : hoja) {
                        if (fila.getRowNum() < 10) continue;

                        Cell celda0 = fila.getCell(0);
                        if (celda0 != null && celda0.getCellTypeEnum() == CellType.STRING) {
                            String texto = celda0.getStringCellValue().trim();
                            Matcher matcher = patronPartida.matcher(texto);
                            if (matcher.find()) {
                                CellStyle style = celda0.getCellStyle();
                                if (style.getFillForegroundColorColor() != null) {
                                    partidaActual = texto;
                                    index_partida++;

                                    Row partidaRow = hojaPartidas.createRow(index_partida);
                                    partidaRow.setHeightInPoints(hojaPartidas.getDefaultRowHeightInPoints() * 1.5f);

                                    Cell cel0 = partidaRow.createCell(0);
                                    cel0.setCellValue(index_partida);
                                    cel0.setCellStyle(estiloConBordes);

                                    Cell cel1 = partidaRow.createCell(1);
                                    cel1.setCellValue(partidaActual);
                                    cel1.setCellStyle(estiloConBordes);

                                    continue;
                                }
                            }
                        }

                        // Procesar funcionario
                        Cell nombre = fila.getCell(6);
                        Cell cargo = fila.getCell(7);
                        Cell monto = fila.getCell(11);

                        if (nombre != null && nombre.getCellTypeEnum() == CellType.STRING && !nombre.getStringCellValue().isEmpty()) {
                            System.out.println("Partida: " + partidaActual);
                            System.out.println("Nombre: " + nombre.getStringCellValue());
                            System.out.println("Cargo: " + (cargo != null ? cargo.getStringCellValue() : "Sin cargo"));
                            System.out.println("Monto: " + (monto != null ? monto.getNumericCellValue() : 0));
                            System.out.println("---");

                            // Aquí puedes escribir también en la hoja[0] para funcionarios si lo deseas
                            index_funcionarios++;
                        }
                    }
                }
            }
        }

        // Finalmente guardamos una sola vez
        try (FileOutputStream fos = new FileOutputStream("Resultados/resultados.xlsx")) {
            workbookOut.write(fos);
            workbookOut.close();
            System.out.println("✅ Archivo resultados.xlsx actualizado.");
        }

        System.out.println(i + " archivos procesados.");
        System.out.println(index_funcionarios + " funcionarios procesados.");
        System.out.println(index_partida + " partidas detectadas.");
    }
}