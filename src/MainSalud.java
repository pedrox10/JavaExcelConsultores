import modelos.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class MainSalud {
    public static void main(String[] args) throws Exception {
        Path carpeta = Paths.get("Consultores/Salud");
        int librosProcesados = 0;
        int index = 0;
        int categoriaPartida = 0;
        // Abrimos el archivo de salida solo una vez
        FileInputStream fisOut = new FileInputStream("Resultados/resultados_salud.xlsx");
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


        try (DirectoryStream<Path> stream = Files.newDirectoryStream(carpeta, "*.xlsx")) {
            for (Path archivo : stream) {
                librosProcesados++;
                try (FileInputStream fis = new FileInputStream(archivo.toFile());
                     Workbook workbook = new XSSFWorkbook(fis)) {
                    Sheet hoja = workbook.getSheet("JULIO");
                    System.out.println("Procesando: " + archivo.getFileName());

                    for (Row fila : hoja) {
                        if (fila.getRowNum() < 1) continue;
                        //Leemos columnas
                        Cell celdaFuente = fila.getCell(2);
                        Cell celdaOrganismo = fila.getCell(3);
                        Cell celdaCarnet = fila.getCell(6);
                        DataFormatter formatter = new DataFormatter();
                        String carnetStr = formatter.formatCellValue(celdaCarnet);
                        Cell celdaPriNombre = fila.getCell(9);
                        Cell celdaSegNombre = fila.getCell(10);
                        Cell celdaPaterno = fila.getCell(11);
                        Cell celdaMaterno = fila.getCell(12);
                        Cell celdaFechaNac = fila.getCell(14);
                        Cell celdaFechaIngreso = fila.getCell(15);
                        Cell celdaGenero = fila.getCell(16);
                        Cell celdaItem = fila.getCell(17);
                        Cell celdaCargo = fila.getCell(18);
                        Cell celdaMonto = fila.getCell(22);

                        if (celdaPriNombre != null && celdaPriNombre.getCellTypeEnum() == CellType.STRING && !celdaPriNombre.getStringCellValue().isEmpty()) {
                            int item = (int) (celdaItem != null ? celdaItem.getNumericCellValue() : 0);
                            if(item == 1)
                                categoriaPartida++;
                            String nombres = celdaPriNombre.getStringCellValue().trim() + " " + celdaSegNombre.getStringCellValue().trim();
                            String paterno = celdaPaterno.getStringCellValue().trim();
                            String materno = celdaMaterno.getStringCellValue().trim();
                            LocalDate fecha = obtenerFecha(celdaFechaNac);
                            String fechaNac = (fecha != null) ? fecha.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) : "";
                            String genero = celdaGenero.getStringCellValue().trim();
                            index++;
                            Funcionario funcionario = new Funcionario(index, paterno, materno, nombres, fechaNac, carnetStr, genero);
                            insertarFuncionario(hojaFuncionarios, funcionario, archivo.getFileName() + "", estiloConBordes);

                            int monto = (int) (celdaMonto != null ? celdaMonto.getNumericCellValue() : 0);

                            int nivel = obtenerNivel(monto);
                            Cargo cargo = new Cargo(index, celdaCargo.getStringCellValue(), nivel, categoriaPartida);
                            insertarCargo(hojaCargos, cargo, estiloConBordes);

                            LocalDate ldIni = obtenerFecha(celdaFechaIngreso);
                            String fechaInicio = (ldIni != null) ? ldIni.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) : "";
                            Contrato contrato = new Contrato(index, item +"", fechaInicio, "", monto);
                            insertarContrato(hojaContratos, contrato, estiloConBordes);

                            String fuente = celdaFuente.getStringCellValue();
                            String organismo = celdaOrganismo.getStringCellValue();
                            PartidaSalud partida = new PartidaSalud(index, "", fuente, organismo, categoriaPartida);
                            insertarPartidaSalud(hojaPartidas, partida, estiloConBordes);
                        }
                    }
                }
            }
        }
        //Finalmente guardamos una sola vez
        try (FileOutputStream fos = new FileOutputStream("Resultados/resultados_salud.xlsx")) {
            workbookOut.write(fos);
            workbookOut.close();
            System.out.println("Archivo resultados.xlsx actualizado.");
        }
        System.out.println(librosProcesados + " archivos procesados.");
        System.out.println(index + " funcionarios procesados.");
        //System.out.println(categoriaPartida + " partidas detectadas.");
    }

    public static void insertarPartidaSalud(Sheet hojaPartidas, PartidaSalud partida, CellStyle estiloConBordes) {
        Row partidaRow = hojaPartidas.createRow(partida.getId());
        partidaRow.setHeightInPoints(hojaPartidas.getDefaultRowHeightInPoints() * 1.5f);
        Cell cel0 = partidaRow.createCell(0);
        cel0.setCellValue(partida.getId());
        cel0.setCellStyle(estiloConBordes);
        Cell cel1 = partidaRow.createCell(1);
        cel1.setCellValue(partida.getNombre());
        cel1.setCellStyle(estiloConBordes);
        Cell cel2 = partidaRow.createCell(2);
        cel2.setCellValue(partida.getFuente());
        cel2.setCellStyle(estiloConBordes);
        Cell cel3 = partidaRow.createCell(3);
        cel3.setCellValue(partida.getOrganismo());
        cel3.setCellStyle(estiloConBordes);
        Cell cel4 = partidaRow.createCell(4);
        cel4.setCellValue(partida.getCategoria());
        cel4.setCellStyle(estiloConBordes);
    }

    public static void insertarFuncionario(Sheet hojaFuncionarios, Funcionario funcionario, String archivoActual, CellStyle estiloConBordes) {
        Row funcionarioRow = hojaFuncionarios.createRow(funcionario.getId());
        funcionarioRow.setHeightInPoints(hojaFuncionarios.getDefaultRowHeightInPoints() * 1.5f);
        Cell cel0 = funcionarioRow.createCell(0);
        cel0.setCellValue(funcionario.getId());
        cel0.setCellStyle(estiloConBordes);
        Cell cel1 = funcionarioRow.createCell(1);
        cel1.setCellValue(funcionario.getNombres());
        cel1.setCellStyle(estiloConBordes);
        Cell cel2 = funcionarioRow.createCell(2);
        cel2.setCellValue(funcionario.getPaterno());
        cel2.setCellStyle(estiloConBordes);
        Cell cel3 = funcionarioRow.createCell(3);
        cel3.setCellValue(funcionario.getMaterno());
        cel3.setCellStyle(estiloConBordes);
        Cell cel4 = funcionarioRow.createCell(4);
        cel4.setCellValue(funcionario.getFechaNac());
        cel4.setCellStyle(estiloConBordes);
        Cell cel5 = funcionarioRow.createCell(5);
        cel5.setCellValue(funcionario.getTextoCi());
        cel5.setCellStyle(estiloConBordes);
        Cell cel6 = funcionarioRow.createCell(6);
        cel6.setCellValue(funcionario.getGenero());
        cel6.setCellStyle(estiloConBordes);
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
        Cell cel2 = funcionarioRow.createCell(2);
        cel2.setCellValue(cargo.getNivel());
        cel2.setCellStyle(estiloConBordes);
        Cell cel3 = funcionarioRow.createCell(3);
        cel3.setCellValue(cargo.getIdPartida());
        cel3.setCellStyle(estiloConBordes);
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
        Cell cel2 = funcionarioRow.createCell(2);
        cel2.setCellValue(contrato.getFechaInicio());
        cel2.setCellStyle(estiloConBordes);
        Cell cel3 = funcionarioRow.createCell(3);
        cel3.setCellValue(contrato.getFechaConclusion());
        cel3.setCellStyle(estiloConBordes);
        Cell cel4 = funcionarioRow.createCell(4);
        cel4.setCellValue(contrato.getMonto());
        cel4.setCellStyle(estiloConBordes);
    }

    public static LocalDate obtenerFecha(Cell celda) {
        if (celda != null && celda.getCellTypeEnum() == CellType.NUMERIC && DateUtil.isCellDateFormatted(celda)) {
            Date fechaJava = celda.getDateCellValue();
            return fechaJava.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        }
        return null;
    }

    public static int obtenerNivel(int monto) {
        switch (monto) {
            case 3435:  return 14;
            case 3553:  return 13;
            case 3761:  return 12;
            case 3958:  return 11;
            case 4379:  return 10;
            case 4586:  return 9;
            case 4786:  return 8;
            case 5200:  return 7;
            case 5707:  return 6;
            case 6290:  return 5;
            case 7295:  return 4;
            case 8186:  return 3;
            case 10494: return 2;
            case 16766: return 1;
            default:    return 0; // No encontrado
        }
    }
}