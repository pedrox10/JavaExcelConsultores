import modelos.Cargo;
import modelos.Contrato;
import modelos.Funcionario;
import modelos.Partida;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Date;
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
                                    insertarPartida(hojaPartidas, partida, estiloConBordes);
                                    continue;
                                }
                            }
                        }
                        // Procesar funcionario
                        Cell minuta = fila.getCell(2);
                        Cell carnet = fila.getCell(3);
                        Cell celdaFecha = fila.getCell(4);
                        Cell nombre = fila.getCell(6);
                        Cell nombre_cargo = fila.getCell(7);
                        Cell celdaMonto = fila.getCell(11);
                        Cell celdaFechaIni = fila.getCell(8);
                        Cell celdaFechaFin = fila.getCell(9);

                        if (nombre != null && nombre.getCellTypeEnum() == CellType.STRING && !nombre.getStringCellValue().isEmpty()) {
                            System.out.println("Nombre: " + nombre.getStringCellValue());
                            String[] arrayNombre = nombre.getStringCellValue().split(" ");
                            String paterno = "";
                            String materno = "";
                            String nombres = "";
                            if(arrayNombre.length > 2){
                                paterno = arrayNombre[0].trim();
                                materno = arrayNombre[1].trim();
                                nombres = String.join(" ", Arrays.copyOfRange(arrayNombre, 2, arrayNombre.length)).trim();
                            } else if (arrayNombre.length == 2) {
                                paterno = arrayNombre[0].trim();
                                nombres = arrayNombre[1].trim();
                            }
                            LocalDate fecha = obtenerFecha(celdaFecha);
                            String fechaNac = (fecha != null) ? fecha.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) : "";
                            int ci = (int) (carnet != null ? carnet.getNumericCellValue() : 0);
                            String respuesta = obtenerGenero(nombres);
                            String genero = "";
                            if(respuesta.equalsIgnoreCase("female")) {
                                genero = "Femenino";
                            } else {
                                if(respuesta.equalsIgnoreCase("male")) {
                                    genero = "Masculino";
                                }
                            }
                            Thread.sleep(500);
                            System.out.println("---");
                            index++;
                            Funcionario funcionario = new Funcionario(index, paterno, materno, nombres, fechaNac, ci, genero);
                            insertarFuncionario(hojaFuncionarios, funcionario, archivo.getFileName() + "", estiloConBordes);
                            int monto = (int) (celdaMonto != null ? celdaMonto.getNumericCellValue() : 0);
                            int nivel = obtenerNivel(monto);
                            Cargo cargo = new Cargo(index, nombre_cargo.getStringCellValue(), nivel, index_partida);
                            insertarCargo(hojaCargos, cargo, estiloConBordes);
                            LocalDate ldIni = obtenerFecha(celdaFechaIni);
                            LocalDate ldFin = obtenerFecha(celdaFechaFin);
                            String fechaInicio = (ldIni != null) ? ldIni.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) : "";
                            String fechaConclusion = (ldFin != null) ? ldFin.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) : "";
                            Contrato contrato = new Contrato(index, minuta.getStringCellValue(), fechaInicio, fechaConclusion, monto);
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

    public static void insertarPartida(Sheet hojaPartidas, Partida partida, CellStyle estiloConBordes) {
        Row partidaRow = hojaPartidas.createRow(partida.getId());
        partidaRow.setHeightInPoints(hojaPartidas.getDefaultRowHeightInPoints() * 1.5f);
        Cell cel0 = partidaRow.createCell(0);
        cel0.setCellValue(partida.getId());
        cel0.setCellStyle(estiloConBordes);
        Cell cel1 = partidaRow.createCell(1);
        cel1.setCellValue(partida.getNombre());
        cel1.setCellStyle(estiloConBordes);
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
        cel5.setCellValue(funcionario.getCi());
        cel5.setCellStyle(estiloConBordes);
        Cell cel6 = funcionarioRow.createCell(6);
        cel6.setCellValue(funcionario.getGenero());
        cel6.setCellStyle(estiloConBordes);
        Cell cel7 = funcionarioRow.createCell(7);
        cel7.setCellValue(archivoActual);
        cel7.setCellStyle(estiloConBordes);
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

    public static String obtenerGenero(String nombres) {
        try {
            String nombresCodificados = URLEncoder.encode(nombres, "UTF-8");
            String apiUrl = "https://api.genderize.io/?name=" + nombresCodificados + "&country_id=BO&apikey=9503d9fe0f7557c932b6662e20c0d09d";
            URL url = new URL(apiUrl);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream(), "UTF-8"));
            StringBuilder response = new StringBuilder();
            String inputLine;
            while ((inputLine = in.readLine()) != null) {
                response.append(inputLine);
            }
            in.close();
            JSONObject json = new JSONObject(response.toString());
            String gender = json.optString("gender", null);
            return gender;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
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