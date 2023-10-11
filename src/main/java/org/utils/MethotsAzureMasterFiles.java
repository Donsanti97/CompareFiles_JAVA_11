package org.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

public class MethotsAzureMasterFiles {

    public static String file1 = /*System.getProperty("user.dir") + */"Documentos\\documents\\initialDocument\\Historico Cartera Comercial.xlsx";
    public static String file2 = /*System.getProperty("user.dir") + */"Documentos\\documents\\finalDocument\\Historico Cartera COMERCIAL por OF.xlsx";

    public static void buscarYListarArchivos(String ubicacion) throws IOException {
        Path ruta = Paths.get(ubicacion);

        if (!Files.exists(ruta)) {
            System.out.println("La ubicación no existe. Creando...");
            Files.createDirectories(ruta);
            System.out.println("Ubicación creada: " + ubicacion);
        } else {
            System.out.println("La ubicación ya existe: " + ubicacion);
            listarArchivosEnCarpeta(ruta);
        }
    }

    public static void listarArchivosEnCarpeta(Path carpeta) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(carpeta)) {
            for (Path archivo : stream) {
                if (Files.isRegularFile(archivo)) {
                    System.out.println("Archivo: " + archivo.getFileName());
                }
            }
        }
    }


    public static String getDocument() {
        // Crea un objeto JFileChooser
        JFileChooser fileChooser = new JFileChooser();

        // Configura el directorio inicial en la carpeta de documentos del usuario
        String rutaDocumentos = System.getProperty("user.home") + File.separator + "Documentos";
        fileChooser.setCurrentDirectory(new File(rutaDocumentos));

        // Filtra para mostrar solo archivos de Excel
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Archivos Excel", "xlsx", "xls"));

        // Muestra el diálogo de selección de archivo
        int resultado = fileChooser.showOpenDialog(null);

        if (resultado == JFileChooser.APPROVE_OPTION) {
            File archivoSeleccionado = fileChooser.getSelectedFile();
            String rutaCompleta = archivoSeleccionado.getAbsolutePath();
            return rutaCompleta;
        } else {
            return null; // Si no se seleccionó ningún archivo, retorna null
        }
    }

    /*-------------------------------------------------------------------------------------------------------------------------------*/
    public static int findSheetIndexInExcelB(String excelAFilePath, String excelBFilePath, String targetSheetName) throws IOException {
        FileInputStream excelAFile = new FileInputStream(excelAFilePath);
        FileInputStream excelBFile = new FileInputStream(excelBFilePath);

        Workbook workbookA = new XSSFWorkbook(excelAFile);
        Workbook workbookB = new XSSFWorkbook(excelBFile);

        int sheetIndexInB = -1;

        for (int i = 0; i < workbookB.getNumberOfSheets(); i++) {
            if (workbookB.getSheetName(i).equals(targetSheetName)) {
                sheetIndexInB = i;
                break;
            }
        }

        List<String> removedSheetNames = new ArrayList<>();

        if (sheetIndexInB != -1) {
            // Elimina las hojas anteriores a la hoja objetivo en Excel B
            for (int i = 0; i < sheetIndexInB; i++) {
                String sheetNameToRemove = workbookB.getSheetName(i);
                removedSheetNames.add(sheetNameToRemove);
            }
        }

        // Cerrar los archivos
        excelAFile.close();
        excelBFile.close();

        return sheetIndexInB;
    }

    public static void runtime() {
        Runtime runtime = Runtime.getRuntime();
        long minRunningMemory = (1024 * 1024);
        if (runtime.freeMemory() < minRunningMemory) {
            System.gc();
        }
    }
    /*---------------------------------------------------------------------------------------------------------------*/

    public static List<String> getWorkSheet(String filePath, int i) {
        List<String> shetNames = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            int numberOfSheets = workbook.getNumberOfSheets();
            ;
            for (int index = i; index < numberOfSheets; index++) {
                Sheet sheet = workbook.getSheetAt(index);
                shetNames.add(sheet.getSheetName());
            }
            workbook.close();
            fis.close();

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return shetNames;
    }

    public static List<Map<String, String>> getValuebyHeader(String excelFilePath, String sheetName) {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = getHeaders(excelFilePath, sheetName);
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String header = headers.get(cellIndex);
                    String value = "";
                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING) {
                            value = cell.getStringCellValue();
                            break;
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            value = String.valueOf(cell.getNumericCellValue());
                        }
                    }
                    rowData.put(header, value);
                }
                data.add(rowData);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }

    public static List<String> getHeaders(String excelFilePath, String sheetName) {
        List<String> headers = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return headers;
    }

    public static List<String> getHeaders(Sheet sheet) {
        List<String> encabezados = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Aquí puedes especificar en qué fila esperas encontrar los encabezados
            // Por ejemplo, si están en la tercera fila (fila índice 2), puedes usar:
            if (row.getRowNum() == 0) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    encabezados.add(obtenerValorCelda(cell));
                }
                break; // Terminamos de buscar encabezados una vez que los encontramos
            }
        }

        return encabezados;
    }

    public static List<String> findValueInColumn(Sheet sheet, int columnaBuscada, String valorBuscado) {
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnaBuscada);
            String valorCelda = obtenerValorCelda(cell);

            if (valorBuscado.equals(valorCelda)) {
                return obtenerValoresFila(row);
            }
        }

        return null; // Valor no encontrado en la columna especificada
    }

    public static List<String> getHeadersMasterfile(Sheet sheet1, Sheet sheet2) throws IOException {
        List<String> headers1 = getHeaders(sheet1);
        String headerFirstFile1 = headers1.get(0);
        List<String> headers2 = getHeaders(sheet2);
        String headerSecondFile = headers2.get(0);

        if (!headerFirstFile1.equals(headerSecondFile)) {
            headers2 = findValueInColumn(sheet1, 0, headers1.get(0));
        }

        return headers2;
    }

    public static List<String> obtenerValoresFila(Row row) {
        List<String> valoresFila = new ArrayList<>();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            valoresFila.add(obtenerValorCelda(cell));
        }
        return valoresFila;
    }

    public static String obtenerValorCelda(Cell cell) {
        String valor = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    valor = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        valor = cell.getDateCellValue().toString();
                    } else {
                        valor = Double.toString(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    valor = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    valor = evaluarFormula(cell);
                    break;
                default:
                    break;
            }
        }
        return valor;
    }

    public static String evaluarFormula(Cell cell) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);
            return cellValue.formatAsString();
        } catch (Exception e) {
            return "";
        }
    }

    public static List<Map<String, String>> obtenerValoresPorFilas(Sheet sheet, List<String> encabezados) {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            Map<String, String> fila = new HashMap<>();
            for (int i = 0; i < encabezados.size() && i < valoresFila.size(); i++) {
                String encabezado = encabezados.get(i);
                String valor = valoresFila.get(i);
                fila.put(encabezado, valor);
            }

            valoresPorFilas.add(fila);
        }

        return valoresPorFilas;
    }

    public static Map<String, String> obtenerValoresPorEncabezado(Sheet sheet, String encabezadoCodCiudad, String encabezadoFecha) {
        Map<String, String> valoresPorCodCiudad = new HashMap<>();

        List<String> encabezados = obtenerValoresFila(sheet.getRow(0)); // Obtener encabezados de la primera fila
        int columnaCodCiudad = -1;
        int columnaFecha = -1;

        // Encontrar las columnas de los encabezados específicos
        for (int i = 0; i < encabezados.size(); i++) {
            String encabezado = encabezados.get(i);
            if (encabezado.equals(encabezadoCodCiudad)) {
                columnaCodCiudad = i;
            }
            if (encabezado.equals(encabezadoFecha)) {
                columnaFecha = i;
            }
        }

        if (columnaCodCiudad == -1 || columnaFecha == -1) {
            return valoresPorCodCiudad; // No se encontraron los encabezados especificados
        }

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String codCiudad = obtenerValorCelda(row.getCell(columnaCodCiudad));
            String valorFecha = obtenerValorCelda(row.getCell(columnaFecha));
            valoresPorCodCiudad.put(codCiudad, valorFecha);
        }

        return valoresPorCodCiudad;
    }


}
