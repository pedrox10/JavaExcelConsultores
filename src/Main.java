import modelos.Cargo;
import modelos.Contrato;
import modelos.Funcionario;
import modelos.Partida;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.regex.*;

public class Main {
    public static void main(String[] args) throws Exception {
        Path carpeta = Paths.get("Consultores");
        int librosProcesados = 0;
        int index = 0;
        int index_partida = 0;
        // Abrimos el archivo de salida solo una vez
        FileInputStream fisOut = new FileInputStream("Resultados/resultados.xlsx");
        Workbook workbookOut = new XSSFWorkbook(fisOut);
        fisOut.close(); // cerrar después de cargar
        // Preparamos hojas y estilos
        Sheet hojaFuncionarios = workbookOut.getSheetAt(0);
        Sheet hojaCargos = workbookOut.getSheetAt(1);
        Sheet hojaContratos = workbookOut.getSheetAt(2);
        Sheet hojaPartidas = workbookOut.getSheetAt(3);
        CellStyle estiloConBordes = workbookOut.createCellStyle();
        estiloConBordes.setBorderTop(BorderStyle.THIN);
        estiloConBordes.setBorderBottom(BorderStyle.THIN);
        estiloConBordes.setBorderLeft(BorderStyle.THIN);
        estiloConBordes.setBorderRight(BorderStyle.THIN);

        Pattern patronPartida = Pattern.compile("^\\d+.*"); // empieza con número

        try (DirectoryStream<Path> stream = Files.newDirectoryStream(carpeta, "*.xlsx")) {
            for (Path archivo : stream) {
                librosProcesados++;
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
                                    Partida partida = new Partida(index_partida, partidaActual);
                                    insertarPartida(hojaPartidas, partida, archivo.getFileName() + "", estiloConBordes);
                                    continue;
                                }
                            }
                        }
                        // Procesar funcionario
                        Cell minuta = fila.getCell(2);
                        Cell nombre = fila.getCell(6);
                        Cell nombre_cargo = fila.getCell(7);
                        Cell monto = fila.getCell(11);

                        if (nombre != null && nombre.getCellTypeEnum() == CellType.STRING && !nombre.getStringCellValue().isEmpty()) {
                            System.out.println("Partida: " + partidaActual);
                            System.out.println("Nombre: " + nombre.getStringCellValue());
                            System.out.println("Contrato: " + minuta.getStringCellValue());
                            System.out.println("Cargo: " + (nombre_cargo != null ? nombre_cargo.getStringCellValue() : "Sin cargo"));
                            System.out.println("Monto: " + (monto != null ? monto.getNumericCellValue() : 0));
                            System.out.println("---");
                            index++;
                            Funcionario funcionario = new Funcionario(index, nombre.getStringCellValue());
                            Cargo cargo = new Cargo(index, nombre_cargo.getStringCellValue());
                            Contrato contrato = new Contrato(index, minuta.getStringCellValue());
                            // Insertando datos en las hojas
                            insertarFuncionario(hojaFuncionarios, funcionario, estiloConBordes);
                            insertarCargo(hojaCargos, cargo, estiloConBordes);
                            insertarContrato(hojaContratos, contrato, estiloConBordes);
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

        System.out.println(librosProcesados + " archivos procesados.");
        System.out.println(index + " funcionarios procesados.");
        System.out.println(index_partida + " partidas detectadas.");
    }

    public static void insertarPartida(Sheet hojaPartidas, Partida partida, String archivoActual, CellStyle estiloConBordes) {
        Row partidaRow = hojaPartidas.createRow(partida.getId());
        partidaRow.setHeightInPoints(hojaPartidas.getDefaultRowHeightInPoints() * 1.5f);
        Cell cel0 = partidaRow.createCell(0);
        cel0.setCellValue(partida.getId());
        cel0.setCellStyle(estiloConBordes);
        Cell cel1 = partidaRow.createCell(1);
        cel1.setCellValue(partida.getNombre());
        cel1.setCellStyle(estiloConBordes);
        Cell cel2 = partidaRow.createCell(2);
        cel2.setCellValue(archivoActual);
        cel2.setCellStyle(estiloConBordes);
    }

    public static void insertarFuncionario(Sheet hojaFuncionarios, Funcionario funcionario, CellStyle estiloConBordes) {
        Row funcionarioRow = hojaFuncionarios.createRow(funcionario.getId());
        funcionarioRow.setHeightInPoints(hojaFuncionarios.getDefaultRowHeightInPoints() * 1.5f);
        Cell cel0 = funcionarioRow.createCell(0);
        cel0.setCellValue(funcionario.getId());
        cel0.setCellStyle(estiloConBordes);
        Cell cel1 = funcionarioRow.createCell(1);
        cel1.setCellValue(funcionario.getNombre());
        cel1.setCellStyle(estiloConBordes);
    }

    public static void insertarCargo(Sheet hojaCargos, Cargo cargo, CellStyle estiloConBordes) {
        Row funcionarioRow = hojaCargos.createRow(cargo.getId());
        funcionarioRow.setHeightInPoints(hojaCargos.getDefaultRowHeightInPoints() * 1.5f);
        Cell cel0 = funcionarioRow.createCell(0);
        cel0.setCellValue(cargo.getId());
        cel0.setCellStyle(estiloConBordes);
        Cell cel1 = funcionarioRow.createCell(1);
        cel1.setCellValue(cargo.getNombre());
        cel1.setCellStyle(estiloConBordes);
    }

    public static void insertarContrato(Sheet hojaContratos, Contrato contrato, CellStyle estiloConBordes) {
        Row funcionarioRow = hojaContratos.createRow(contrato.getId());
        funcionarioRow.setHeightInPoints(hojaContratos.getDefaultRowHeightInPoints() * 1.5f);
        Cell cel0 = funcionarioRow.createCell(0);
        cel0.setCellValue(contrato.getId());
        cel0.setCellStyle(estiloConBordes);
        Cell cel1 = funcionarioRow.createCell(1);
        cel1.setCellValue(contrato.getMinuta());
        cel1.setCellStyle(estiloConBordes);
    }
}