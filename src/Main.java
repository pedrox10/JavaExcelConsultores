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
                            String genero = obtenerGenero(nombres);
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

    /*public static String obtenerGenero(String nombres) {
        try {
            String nombresCodificados = URLEncoder.encode(nombres, "UTF-8");
            String apiUrl = "https://api.genderize.io/?name=" + nombresCodificados + "&country_id=BO&apikey=14674a6ff426d39cd8b09d21443952c4";
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
    }*/

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

    public static String obtenerGenero(String nombre) {
        if (nombre == null || nombre.trim().isEmpty()) {
            return "";
        }
        nombre = nombre.trim().toUpperCase();

        switch (nombre) {
            // Masculinos
            case "OLIVER FERNANDO":
            case "ROMER FREDDY":
            case "ERIK ANDRE":
            case "CLEMENTE":
            case "ANGEL JHEYSON":
            case "JOSE CARLOS":
            case "JUAN CARLOS":
            case "IVAN":
            case "GROVER":
            case "HENRY":
            case "GONZALO":
            case "CARLOS BORIS":
            case "JHOEL ALEJANDRO":
            case "MARCO VINICIO":
            case "RAMIRO":
            case "PABLO":
            case "CONSTANCIO":
            case "ARIEL":
            case "BENJI":
            case "SALVADOR KARIM":
            case "FERNANDO":
            case "BRAYAN FERNANDO":
            case "JOSE BELTRAN":
            case "RAUL":
            case "ALBERTO":
            case "ORLANDO DANNY":
            case "DANIEL":
            case "SAMUEL ALEX":
            case "BRIAN DANNY":
            case "FABIAN JUAN":
            case "MARCELO":
            case "MIJAIL":
            case "JOSE NELSON":
                return "Masculino";

            // Femeninos
            case "KATERIN LINET":
            case "CIELO STEFANIA":
            case "ELSA":
            case "MARIA LUISA":
            case "YOLANDA":
            case "NORMA":
            case "ELIZABET XIMENA":
            case "LIZETH":
            case "JUDITH CELENA":
            case "ZAIDA":
            case "ELVIA":
            case "FABIOLA":
            case "PAOLA":
            case "DANIELA":
            case "MARGARITA":
            case "BEATRIZ":
            case "VIRGINIA":
            case "ELIZABETH SUSANA":
            case "NATALY":
            case "VANESSA":
            case "MARIBEL ROCIO":
            case "AMANDA":
            case "NANCY":
            case "GIMENA":
            case "DANIA":
            case "ASUNTA":
            case "JHOSELYN":
            case "MICAELA VANESA":
            case "VILMA":
            case "NOELIA CLAUDIA":
            case "TANIA ROSARIO":
            case "LILIANA PATRICIA":
            case "GUETSI MAYA":
            case "NIKA YOLANDA":
            case "LEICY ALEJANDRA":
            case "RELINDA":
            case "MARIA LIZETH":
            case "MAURA":
            case "EVELIN WENDY":
            case "MARIZOL":
            case "YAMILET MILAYDA":
            case "ROSSYZELA":
            case "LEYDI":
            case "LUZ MARIAN":
            case "KARLA LORENA":
            case "LINDA VERONICA":
            case "MARIA ELVIA":
            case "KAREN OLIVIA":
            case "DANIDZA":
            case "MARIBEL":
            case "ANAI":
            case "MARTHA":
            case "AIDEE":
            case "CARLA ELIZABETH":
            case "CARMEN LIZZETTE":
            case "SILVIA EUGENIA":
            case "LIZ":
            case "ELIZABETH":
            case "CLAUDIA":
            case "NELY":
            case "VANIA":
            case "ADELA":
            case "JULIETA":
            case "MARTHA LILIANA":
            case "CATHERINE PATRICIA":
            case "JIMENA":
            case "ANDREA":
            case "CRISTINA":
            case "MARIA LEYDI":
            case "MARISOL":
            case "MARIANA YOLANDA":
            case "JULIETA CRISTINA":
            case "MARITZA ZOBEIDA":
            case "MARIA LOURDES":
            case "MILEYKA":
            case "ANGELA":
            case "IBETH FABIOLA":
            case "GUILLERMINA":
            case "CINTHIA":
            case "EVELYN DAYANA":
            case "TECHY":
            case "WENDY":
            case "LIDIA":
            case "ARMINDA":
            case "NICCOL":
            case "JULIA":
            case "GERALDINE KIMBERLY":
            case "FIDELIA":
            case "REBECA":
            case "ZULMA ANTONIA":
            case "MARINA":
            case "ANEYDA":
            case "ANDREA HELEN":
            case "ELIANA VERONICA":
            case "DAYANA DEL ROSARIO":
            case "ADRIANA JACKELINE":
            case "ADRIANA YASMIN":
            case "JESICA KARINA":
            case "JENNIFER":
            case "ELIANA":
            case "ANA MARIA":
            case "SILENE":
            case "GEORGINA":
            case "JOANA":
            case "MIRIAN":
            case "ROSSE MERY":
            case "SELENE":
            case "CLEIDY":
            case "MARIA NEISA":
            case "DIANECA":
            case "NICOLE":
            case "MARGOT TOMASA":
            case "BRENDA ANDREA":
            case "NATALI GEOVANNA":
            case "HAZEL":
            case "ANDREA ALEJANDRA":
            case "ESTHER":
            case "MARY LIZETH":
            case "DEYSI":
            case "NOEMI":
                return "Femenino";

            default:
                return "";
        }
    }
}