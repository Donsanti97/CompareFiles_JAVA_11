package org.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MethotsAzureMasterFiles {

    public static final String SPECIAL_CHAR = " -X- ";
    public static List<String> errores = new ArrayList<>();
    public  static List<String> coincidencias = new ArrayList<>();



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
        fileChooser.setFileFilter(new FileNameExtensionFilter("Archivos Excel", "xlsx", "xls"));

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

    public static String getDirectory() {
        // Crea un objeto JFileChooser
        JFileChooser fileChooser = new JFileChooser();

        // Configura el directorio inicial en la carpeta de documentos del usuario
        String rutaDocumentos = System.getProperty("user.home")/* + File.separator + "Documentos"*/;
        fileChooser.setCurrentDirectory(new File(rutaDocumentos));

        // Filtra para mostrar solo archivos de Excel
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

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
        long minRunningMemory = (8L * 1024L * 1024L * 1024L);
        if (runtime.freeMemory() < minRunningMemory) {
            System.gc();
        }
    }
    /*---------------------------------------------------------------------------------------------------------------*/

    public static List<String> getWorkSheet(String filePath, int i) {
        List<String> shetNames = new ArrayList<>();
        try {
            Workbook workbook = WorkbookFactory.create(new File(filePath));
            int numberOfSheets = workbook.getNumberOfSheets();

            for (int index = i; index < numberOfSheets; index++) {
                Sheet sheet = workbook.getSheetAt(index);
                shetNames.add(sheet.getSheetName());
            }
            workbook.close();

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
                    encabezados.add(obtenerValorVisibleCelda(cell));
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
            String valorCelda = obtenerValorVisibleCelda(cell);

            if (valorBuscado.equals(valorCelda)) {
                return obtenerValoresFila(row);
            }
        }

        return null; // Valor no encontrado en la columna especificada
    }

    public static List<String> headersRow(String filePath, String sheetName, String targetHeader) {
        try (Workbook workbook = WorkbookFactory.create(new File(filePath))) {
            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet != null) {
                for (Row row : sheet) {
                    Cell cell = row.getCell(0);

                    if (cell != null) {
                        String cellValue = obtenerValorVisibleCelda(cell);
                        if (targetHeader.equalsIgnoreCase(cellValue)) {
                            int rowNum = row.getRowNum() + 1;

                        }
                    }
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        return null;
    }

    public static int findHeaderRow(String filePath, String sheetName, String targetHeader) {
        try (Workbook workbook = WorkbookFactory.create(new File(filePath))) {
            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet != null) {
                for (Row row : sheet) {
                    Cell cell = row.getCell(0); // Primera columna

                    if (cell != null) {
                        String cellValue = cell.getStringCellValue();
                        if (targetHeader.equalsIgnoreCase(cellValue)) {
                            return row.getRowNum() + 1; // Se suma 1 porque las filas se cuentan desde 0
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return -1; // Retornar -1 si no se encuentra el encabezado
    }

    public static List<String> getHeadersMasterfile(Sheet sheet1, Sheet sheet2, String seleccion) throws IOException {
        List<String> headers1 = getHeaders(sheet1);
        String headerFirstFile1 = headers1.get(0);
        List<String> headers2 = getHeaders(sheet2);
        String headerSecondFile = headers2.get(0);

        if (!headerFirstFile1.equals(headerSecondFile)) {
            headers2 = findValueInColumn(sheet2, 0, seleccion);
        }

        return headers2;
    }



    public static List<String> getHeadersMasterfile(Sheet sheet1, Sheet sheet2) throws IOException {
        List<String> headers1 = getHeaders(sheet1);
        //String headerFirstFile1 = headers1.get(0);
        List<String> headers2 = getHeaders(sheet2);
        String headerSecondFile = headers2.get(0);
        for (String headerFirstFile1: headers1){
            if (!headerFirstFile1.equals(headerSecondFile)) {
                headers2 = findValueInColumn(sheet2, 0, headerFirstFile1);
            }
        }



        return headers2;
    }

    public static List<String> obtenerValoresFila(Row row) {
        List<String> valoresFila = new ArrayList<>();
        //Iterator<Cell> cellIterator = row.cellIterator();
        int index = row.getLastCellNum();
        for (int i = 0; i < index; i++) {
            Cell cells = row.getCell(i);
            String value = obtenerValorVisibleCelda(cells);
            if (value == null || value.isBlank() || value.isEmpty() || value.equals("null")){
                value = "0";
            }
            valoresFila.add(value);
        }
        /*while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String value = obtenerValorVisibleCelda(cell);
            if (Objects.equals(value, "null") || value == null || value.isEmpty()){
                value = "0";
            }
            valoresFila.add(value);
        }*/
        return valoresFila;
    }

    public static List<String> obtenerValoresFilaAzure(Row row) {
        List<String> valoresFila = new ArrayList<>();
        int index = row.getLastCellNum();
         String values;
        for (int i = 0; i < index; i++) {
            Cell cells = row.getCell(i);
            values = obtenerValorVisibleCelda(cells);
            if (values == null || values.equals("null") || values.isEmpty() || values.isBlank()){
                values = "0";
            }
            valoresFila.add(values);

        }

        return valoresFila;
    }

    public static String obtenerValorCelda(Cell cell) {
        String valor = "";
        if (cell != null) {
            try {
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getStringCellValue() + ", CELL: " + cell);
                        valor = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            //valor = cell.getDateCellValue().toString();
                            System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getNumericCellValue() + ", CELL: " + cell);
                            String formatDate = cell.getDateCellValue().toString();
                            SimpleDateFormat formatoEntrada = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                            Date date = formatoEntrada.parse(formatDate);
                            SimpleDateFormat formatoSalida = new SimpleDateFormat("dd/MM/yyyy");
                            //valor = formatoSalida.format(date);
                            valor = cell.toString();
                            System.out.println("VALOR1 " + valor);
                        } else {
                            //valor = Date.toString(cell.getDateCellValue()/*NumericCellValue()*/);
                            System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getNumericCellValue() + ", CELL: " + cell);
                            valor = cell.getStringCellValue();
                            System.out.println("VALOR3 " + valor);
                        }
                        break;
                    case BOOLEAN:
                        System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getBooleanCellValue() + ", CELL: " + cell);
                        valor = Boolean.toString(cell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        System.out.println("CELLTYPE " + cell.getCellType() + ": " + cell.getCellFormula().toString() + ", CELL: " + cell + " CELLADD: " + cell.getAddress());
                        /*System.err.println("Formato fecha no valido."+ cell.getCellFormula() +" Encabezado "+ cell.getSheet().getSheetName() +" Posición: "+ cell.getAddress() +" puede contener formula o valor cadena de caracteres");
                        valor = evaluarFormula(cell);
                        System.out.println("VALORF: " + valor);
                        FunctionsApachePoi.waitSeconds(20);
                        System.exit(1);*/
                        valor = obtenerValorCeldaString(cell);

                        //break;
                    default:
                        /*valor = obtenerValorCeldaString(cell);*/
                        break;
                }

            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        return valor;
    }

    public static String obtenerValorVisibleCelda(Cell cell) {
        try {
            DataFormatter dataFormatter = new DataFormatter();
            String valor = "";

            // Verificar el tipo de celda
            switch (cell.getCellType()) {
                case STRING:
                    valor = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        valor = dataFormatter.formatCellValue(cell);
                    } else {
                        double numericValue = cell.getNumericCellValue();
                        String dataFormatString = cell.getCellStyle().getDataFormatString();

                        if (numericValue >= -99.99 && numericValue <= 99.99) {
                            if (numericValue == 0) {
                                valor = dataFormatter.formatRawCellContents(cell.getNumericCellValue(), cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString());

                            } else {
                                boolean isTwoDigitsOrLess = Math.abs(numericValue) < 100 && Math.abs(numericValue) % 1 != 0;
                                if (isTwoDigitsOrLess) {
                                    valor = String.format("%.2f%%", numericValue/* / 100*/);
                                } else {
                                    valor = String.valueOf(numericValue);
                                }
                            }
                        } else {
                            valor = dataFormatter.formatRawCellContents(cell.getNumericCellValue(), cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString());

                        }
                    }
                    break;
                case BOOLEAN:
                    valor = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    valor = evaluarFormulas(cell);

                case BLANK:
                case _NONE:
                case ERROR:
                default:
                    valor = /*dataFormatter.formatCellValue(cell)*/"0";
            }

            if (cell.getCellType() == null){
                valor = "0";
            }

            return valor;
        } catch (Exception e) {
            return "";
        }
    }

    public static String evaluarFormulas(Cell cell) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);

            if (cellValue == null){
                return "0";
            }
            switch (cellValue.getCellType()) {
                case STRING:
                    return cellValue.getStringValue();
                case NUMERIC:
                    return String.valueOf(cellValue.getNumberValue());
                case BOOLEAN:
                    return String.valueOf(cellValue.getBooleanValue());
                case ERROR:
                    return "Error: " + cellValue.getErrorValue();
                case BLANK:
                case _NONE:
                    return "0";
                default:
                    return cellValue.formatAsString();
            }
            /*if (cellValue.getCellType() == CellType.NUMERIC) {
                double valor = cellValue.getNumberValue();
                //System.out.println("El valor de la fórmula en A5 es: " + valor);
                return Double.toString(valor);
            } else if (cellValue.getCellType() == CellType.STRING) {
                String valor = cellValue.getStringValue();
                //System.out.println("El valor de la fórmula en A5 es: " + valor);
                return valor;
            } else {
                return cellValue.formatAsString();
            }*/
        } catch (Exception e) {
            return "";
        }
    }

    public static String formatearNumeroConPuntos(String numero) {
        // Convertir el número a cadena y verificar si ya tiene puntos de separación de miles
        //String numeroStr = String.valueOf(numero);
        if (!numero.contains(".")) {
            // Si no tiene puntos, agregarlos
            return agregarPuntosSeparadores(numero);
        } else {
            // Si ya tiene puntos, devolver la cadena original
            return numero;
        }
    }

    public static String agregarPuntosSeparadores(String numeroStr) {
        // Dividir la parte entera y la parte decimal (si existe)
        String[] partes = numeroStr.split("\\.");
        String parteEntera = partes[0];
        String parteDecimal = partes.length > 1 ? "." + partes[1] : "";

        // Agregar puntos de separación de miles a la parte entera
        StringBuilder resultado = new StringBuilder();
        int contador = 0;
        for (int i = parteEntera.length() - 1; i >= 0; i--) {
            resultado.insert(0, parteEntera.charAt(i));
            contador++;
            if (contador == 3 && i > 0) {
                resultado.insert(0, ".");
                contador = 0;
            }
        }

        // Combinar la parte entera formateada y la parte decimal
        return resultado.toString() + parteDecimal;
    }

    public static String obtenerValorCeldaString(Cell cell) {
        try {
            DataFormatter dataFormatter = new DataFormatter();
            String valor = dataFormatter.formatCellValue(cell);
            return valor;
        } catch (Exception e) {
            return "";
        }
    }

    public static String evaluarFormula(Cell cell) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);

            if (cellValue.getCellType() == CellType.FORMULA) {
                // Si la celda contiene una fórmula, obtén su valor calculado
                if (cellValue.getCellType() == CellType.NUMERIC) {
                    double valor = cellValue.getNumberValue();
                    return Double.toString(valor);
                } else if (cellValue.getCellType() == CellType.STRING) {
                    return cellValue.getStringValue();
                }
            } else {
                // Si no es una fórmula, obtén el valor directo de la celda
                System.out.println("NO ES FORMULA");
                DataFormatter dataFormatter = new DataFormatter();
                String valor = dataFormatter.formatCellValue(cell);
                return valor;
            }
        } catch (Exception e) {
            return "";
        }

        return ""; // Valor por defecto si no se pudo obtener el valor de la fórmula
    }

    public static List<Map<String, String>> createMapList(List<Map<String, String>> originalList, String keyHeader, String valueHeader) {
        List<Map<String, String>> mapList = new ArrayList<>();

        try {
            for (Map<String, String> originalMap : originalList) {
                String key = originalMap.get(keyHeader);
                String value = originalMap.get(valueHeader);

                //System.out.println("AQUÍ LLENA EL MAP_LIST. \n KEY: " + key + ", VALUE: " + value);
                Map<String, String> newMap = new HashMap<>();
                Map<String, String> errorMap = new HashMap<>();

                //System.out.println( "KEY: " + key + ", VALUE: " + value );

                if (!key.equals(keyHeader) && !value.equals(valueHeader)) {
                    if (key == "null") {
                        //System.out.println("ENTRA AL CONDICIONAL NULL");
                        errorMap.put(key, value);
                    } else {
                        newMap.put(key, value);
                    }
                } else {
                    errorMap.put(key, value);
                }



                mapList.add(newMap);
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return mapList;
    }

    public static String ponerDecimales(String value){
        String newValue = "";
        try {
            if (value.isEmpty() || value.equals("-")){
                value = "0";
            }
            if (!value.contains(".") || value.equals("0")) {
                if (value.contains("-")){ value = value.replace("-", "");}
                if (value.contains("%")){ value = value.replace("%", "");}
                value = value.replace(",", ".");
                double number = Double.parseDouble(value);
                DecimalFormat df = new DecimalFormat("#");//"#,##0.00"
                newValue = df.format(number);
                //System.err.println("VALUE: " + newValue);

                return newValue;
            } else {
                /**/
                if (value.contains("%")){ value = value.replace("%", "");}
                if (value.length() > 6){
                    value = value.replace(".", "").replace(",", ".");
                }
                double number = Double.parseDouble(value);
                DecimalFormat df = new DecimalFormat("#");//"#,##0.00"
                value = df.format(number);
                /**/
                return value;
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static List<Map<String, String>> obtenerValoresPorFilas(Workbook workbook1, Workbook workbook2, String sheetName1, String sheetName2, String header1, String header2, String firstMHeader) throws IOException {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();
        Sheet sheet1 = workbook1.getSheet(sheetName1);
        Sheet sheet2 = workbook2.getSheet(sheetName2);

        List<String> encabezados = getHeadersMasterfile(sheet1, sheet2, firstMHeader);

        int indexHeader1 = encabezados.indexOf(header1);
        int indexHeader2 = encabezados.indexOf(header2);

        int count = 0;
        int i = 0;
        int rowsPerBatch = 5000;

        if (indexHeader1 == -1 || indexHeader2 == -1) {
            System.err.println("Los encabezados no se encontraron en la hoja " + sheetName2);
            valoresPorFilas = null;
            return valoresPorFilas;
            //throw new IllegalArgumentException("Los encabezados especificados no se encontraron en la hoja.");
        }

        Iterator<Row> rowIterator = sheet2.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            Map<String, String> fila = new HashMap<>();

            try {
                while (valoresFila.size() != encabezados.size()){
                    valoresFila.add("0");
                }
                if (indexHeader1 >= 0 && indexHeader1 <= valoresFila.size() &&
                        indexHeader2 >= 0 && indexHeader2 <= valoresFila.size()) {
                    fila.put(header1, valoresFila.get(indexHeader1));
                    fila.put(header2, valoresFila.get(indexHeader2));
                    count++;
                } else {
                    System.err.println("En la fila [" + row.getRowNum() + "] no se encuentran los datos completos. El valor no puede ser nulo" +
                            "\n Por favor rellene con [0] o con [NA] según el campo que falte numérico o caracteres respectivamente");
                    i++;
                }
                if (count % rowsPerBatch == 0) {
                    runtime();
                    Thread.sleep(200);
                }
                System.err.println();
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
            valoresPorFilas.add(fila);

        }
        int total = count + i;
        System.err.println("NUMERO DE FILAS VALIDADAS: " + total +
                "\n NUMERO DE FILAS NO ANALIZADAS: " + i +
                "\n NUMERO DE FILAS ANALIZADAS: " + count);
        if (total == i){
            errorMessage("No es posible continuar con el análisis, la cantidad de información incompleta es demasiada." +
                    "\n Por favor verifique las indicaciones anteriores.");
            return null;
        }else {
            return valoresPorFilas;
        }
    }

    public static List<Map<String, String>> obtenerValoresPorFilas(Workbook workbook1, Workbook workbook2, String sheetName1, String sheetName2) throws IOException {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();
        Sheet sheet1 = workbook1.getSheet(sheetName1);
        Sheet sheet2 = workbook2.getSheet(sheetName2);

        List<String> encabezados = getHeadersMasterfile(sheet1, sheet2);

        Iterator<Row> rowIterator = sheet1.iterator();
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

    public static List<Map<String, String>> obtenerValoresPorFilas(Sheet sheet, Sheet sheet2) throws IOException {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();
        List<String> encabezados = getHeadersMasterfile(sheet, sheet2);

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

    public static List<Map<String, String>> obtenerValoresPorFilas(Sheet sheet, List<String> encabezados, String header1, String header2) throws IllegalArgumentException {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();

        int indexHeader1 = encabezados.indexOf(header1);
        int indexHeader2 = encabezados.indexOf(header2);

        int count = 0;
        int i = 0;
        int rowsPerBatch = 5000;

        if (indexHeader1 == -1 || indexHeader2 == -1){
            System.err.println("Los encabezados no se encontraron en la hoja " + sheet.getSheetName());
            valoresPorFilas = null;
            return valoresPorFilas;
            //throw  new IllegalArgumentException("Los encabezados especificados no se encontraron en la hoja.");

        }

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFilaAzure(row);

            Map<String, String> fila = new HashMap<>();

            try {
                while (valoresFila.size() != encabezados.size()){
                    valoresFila.add("0");
                }
                if (indexHeader1 >= 0 && indexHeader1 <= valoresFila.size() &&
                        indexHeader2 >= 0 && indexHeader2 <= valoresFila.size()) {
                    fila.put(header1, valoresFila.get(indexHeader1));
                    fila.put(header2, valoresFila.get(indexHeader2));

                    count++;
                }else {
                    i++;
                }

                if (count % rowsPerBatch == 0) {
                    runtime();
                    Thread.sleep(200);
                }
                System.out.println();

            } catch (Exception e) {
                throw new RuntimeException(e);
            }

            valoresPorFilas.add(fila);
        }
        int total = count + i;
        System.err.println("NUMERO DE FILAS VALIDADAS: " + total +
                "\n NUMERO DE FILAS NO ANALIZADAS: " + i +
                "\n NUMERO DE FILAS ANALIZADAS: " + count);
        if (total == i){
            errorMessage("No es posible continuar con el análisis, la cantidad de información incompleta es demasiada." +
                    "\n Por favor verifique las indicaciones anteriores.");
            return null;
        }else {
            return valoresPorFilas;
        }
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
            String codCiudad = obtenerValorVisibleCelda(row.getCell(columnaCodCiudad));
            String valorFecha = obtenerValorVisibleCelda(row.getCell(columnaFecha));
            valoresPorCodCiudad.put(codCiudad, valorFecha);
        }

        return valoresPorCodCiudad;
    }

    public static void errorMessage(String mensaje) {
        JLabel label = new JLabel("<html><font color='red'>" + mensaje + "</font></html>");
        label.setFont(new Font("Arial", Font.PLAIN, 14)); // Puedes ajustar la fuente según tus preferencias

        JOptionPane.showMessageDialog(null, label, "Error", JOptionPane.ERROR_MESSAGE);
    }

    public static void waitSeconds(int seconds) {
        try {
            Thread.sleep((seconds * 1000L));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /*public static String mostrarMenu(List<String> opciones) {

        opciones.add(0, "Ninguno");

        JFrame frame = new JFrame("Menú de Opciones");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JComboBox<String> comboBox = new JComboBox<>(opciones.toArray(new String[0]));
        comboBox.setSelectedIndex(0);

        JButton button = new JButton("Seleccionar");


        ActionListener actionListener = e -> frame.dispose();

        button.addActionListener(actionListener);

        JPanel panel = new JPanel();
        panel.add(comboBox);
        panel.add(button);

        frame.add(panel);
        frame.setSize(300, 100);
        frame.setVisible(true);

        while (frame.isVisible()) {
            // Esperar hasta que la ventana se cierre
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        return comboBox.getSelectedItem().toString();
    }*/

    public static String mostrarMenu(List<String> opciones) {
        List<String> opcionesConNinguno = new ArrayList<>(opciones);
        opcionesConNinguno.add(0, "Ninguno");

        JFrame frame = new JFrame("Menú de Opciones");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JComboBox<String> comboBox = new JComboBox<>(opcionesConNinguno.toArray(new String[0]));
        comboBox.setSelectedIndex(0);

        JButton button = new JButton("Seleccionar");

        ActionListener actionListener = e -> frame.dispose();

        button.addActionListener(actionListener);

        JPanel panel = new JPanel();
        panel.add(comboBox);
        panel.add(button);

        frame.add(panel);
        frame.setSize(300, 100);
        frame.setVisible(true);

        while (frame.isVisible()) {
            // Esperar hasta que la ventana se cierre
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        return comboBox.getSelectedItem().toString();
    }


    public static String showYesNoDialog(String message) {
        JFrame frame = new JFrame();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // El mensaje que se mostrará en el cuadro de diálogo
        String[] options = {"SI", "NO"};

        // Mostrar el cuadro de diálogo con los botones "Sí" y "No"
        int choice = JOptionPane.showOptionDialog(
                frame,
                message,
                "Confirmación",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                options,
                options[0]
        );

        frame.dispose();

        // Retorna la opción seleccionada como String
        return (choice == JOptionPane.YES_OPTION) ? "SI" : "NO";
    }

    public static List<String> getHeadersN(Sheet sheet) {
        List<String> columnNames = new ArrayList<>();
        Row row = sheet.getRow(0);
        try {
            System.out.println("PROCESANDO CAMPOS...");
            for (Iterator<Cell> it = row.cellIterator(); it.hasNext(); ) {
                Cell cell = it.next();
                columnNames.add(obtenerValorVisibleCelda(cell));
                //System.out.println(obtenerValorVisibleCelda(cell));
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        return columnNames;
    }

    /*public static List<String> createDualDropDownListsAndReturnSelectedValues(List<String> list1, List<String> list2) {
        List<String> selectedValues = new ArrayList<>();

        JFrame frame = new JFrame("SELECCIÓN DE HOJAS");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 200);
        frame.setLayout(new FlowLayout());

        JComboBox<String> dropdown1 = new JComboBox<>(list1.toArray(new String[0]));
        JComboBox<String> dropdown2 = new JComboBox<>(list2.toArray(new String[0]));
        JButton addButton = new JButton("Agregar Selecciones");

        frame.add(dropdown1);
        frame.add(dropdown2);
        frame.add(addButton);

        // Panel para contener las selecciones y checkboxes
        JPanel selectionsPanel = new JPanel(new GridLayout(0, 2));
        JScrollPane scrollPane = new JScrollPane(selectionsPanel);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        frame.add(selectionsPanel);

        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String selectedValue1 = (String) dropdown1.getSelectedItem();
                String selectedValue2 = (String) dropdown2.getSelectedItem();

                if (selectedValue1 != null && selectedValue2 != null) {
                    String combinedSelection = selectedValue1 + SPECIAL_CHAR + selectedValue2;
                    selectedValues.add(combinedSelection);

                    // Crear checkbox para la selección recién agregada
                    JCheckBox checkBox = new JCheckBox(combinedSelection);
                    selectionsPanel.add(checkBox);

                    // Eliminar elementos seleccionados de los desplegables
                    list1.remove(selectedValue1);
                    list2.remove(selectedValue2);

                    // Actualizar los modelos de los desplegables
                    dropdown1.setModel(new DefaultComboBoxModel<>(list1.toArray(new String[0])));
                    dropdown2.setModel(new DefaultComboBoxModel<>(list2.toArray(new String[0])));

                    System.out.println("Elementos agregados: " + combinedSelection);

                    frame.revalidate();
                    frame.repaint();
                } else {
                    // Puedes mostrar un mensaje de error si ambos elementos no están seleccionados
                    JOptionPane.showMessageDialog(frame, "Selecciona un elemento de cada lista", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Botón para eliminar selecciones marcadas
        JButton removeButton = new JButton("Eliminar Selecciones");
        frame.add(removeButton);

        removeButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Eliminar selecciones marcadas
                for (Component component : selectionsPanel.getComponents()) {
                    if (component instanceof JCheckBox) {
                        JCheckBox checkBox = (JCheckBox) component;
                        if (checkBox.isSelected()) {
                            selectedValues.remove(checkBox.getText());

                            // Recuperar elementos eliminados a los desplegables
                            String[] parts = checkBox.getText().split(SPECIAL_CHAR);
                            if (!list1.contains(parts[0])) {
                                list1.add(parts[0]);
                            }
                            if (!list2.contains(parts[1])) {
                                list2.add(parts[1]);
                            }

                            // Actualizar los modelos de los desplegables
                            dropdown1.setModel(new DefaultComboBoxModel<>(list1.toArray(new String[0])));
                            dropdown2.setModel(new DefaultComboBoxModel<>(list2.toArray(new String[0])));

                            selectionsPanel.remove(checkBox);
                        }
                    }
                }

                frame.revalidate();
                frame.repaint();
            }
        });

        // Botón para terminar el proceso de selección
        JButton finishButton = new JButton("Terminar Selección");
        frame.add(finishButton);

        finishButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Puedes realizar acciones finales aquí, por ejemplo, cerrar la aplicación
                frame.dispose();
            }
        });

        frame.setVisible(true);

        // Esperar hasta que se cierre la ventana
        while (frame.isVisible()) {
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        return selectedValues;
    }*/

    public static List<String> createDualDropDownListsAndReturnSelectedValues(List<String> list1, List<String> list2) {
        List<String> selectedValues = new ArrayList<>();

        JFrame frame = new JFrame("SELECCIÓN DE HOJAS");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 300);
        frame.setLayout(new FlowLayout());

        JComboBox<String> dropdown1 = new JComboBox<>(list1.toArray(new String[0]));
        JComboBox<String> dropdown2 = new JComboBox<>(list2.toArray(new String[0]));
        JButton addButton = new JButton("Agregar Selecciones");

        frame.add(dropdown1);
        frame.add(dropdown2);
        frame.add(addButton);

        DefaultListModel<String> listModel = new DefaultListModel<>();
        JList<String> selectionsList = new JList<>(listModel);
        JScrollPane scrollPane = new JScrollPane(selectionsList);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        frame.add(scrollPane);

        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String selectedValue1 = (String) dropdown1.getSelectedItem();
                String selectedValue2 = (String) dropdown2.getSelectedItem();

                if (selectedValue1 != null && selectedValue2 != null) {
                    String combinedSelection = selectedValue1 + SPECIAL_CHAR + selectedValue2;
                    selectedValues.add(combinedSelection);

                    listModel.addElement(combinedSelection);

                    // Eliminar elementos seleccionados de los desplegables
                    list1.remove(selectedValue1);
                    list2.remove(selectedValue2);

                    // Actualizar los modelos de los desplegables
                    dropdown1.setModel(new DefaultComboBoxModel<>(list1.toArray(new String[0])));
                    dropdown2.setModel(new DefaultComboBoxModel<>(list2.toArray(new String[0])));

                    System.out.println("Elementos agregados: " + combinedSelection);
                } else {
                    // Puedes mostrar un mensaje de error si ambos elementos no están seleccionados
                    JOptionPane.showMessageDialog(frame, "Selecciona un elemento de cada lista", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Botón para eliminar selecciones marcadas
        JButton removeButton = new JButton("Eliminar Selecciones");
        frame.add(removeButton);

        removeButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Eliminar selecciones marcadas
                int[] selectedIndices = selectionsList.getSelectedIndices();
                for (int i = selectedIndices.length - 1; i >= 0; i--) {
                    String removedValue = listModel.getElementAt(selectedIndices[i]);
                    selectedValues.remove(removedValue);

                    // Recuperar elementos eliminados a los desplegables
                    String[] parts = removedValue.split(SPECIAL_CHAR);
                    if (!list1.contains(parts[0])) {
                        list1.add(parts[0]);
                    }
                    if (!list2.contains(parts[1])) {
                        list2.add(parts[1]);
                    }

                    listModel.removeElementAt(selectedIndices[i]);
                }

                // Actualizar los modelos de los desplegables
                dropdown1.setModel(new DefaultComboBoxModel<>(list1.toArray(new String[0])));
                dropdown2.setModel(new DefaultComboBoxModel<>(list2.toArray(new String[0])));
            }
        });

        // Botón para terminar el proceso de selección
        JButton finishButton = new JButton("Terminar Selección");
        frame.add(finishButton);

        finishButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // Puedes realizar acciones finales aquí, por ejemplo, cerrar la aplicación
                frame.dispose();
            }
        });

        frame.setVisible(true);

        // Esperar hasta que se cierre la ventana
        while (frame.isVisible()) {
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        return selectedValues;
    }


    public static void logWinsToFile(String filePath, List<String> messages) {
        writeExcelFile(filePath, messages, "messages");
    }

    public static void logErrorsToFile(String filePath, List<String> errors) {
        writeExcelFile(filePath, errors, "errors");
    }

    public static void writeExcelFile(String filePath, List<String> messages, String folderName) {
        // Obtener el nombre del archivo sin la extensión
        String fileName = new File(filePath).getName();
        String folderPath = filePath.replace(fileName, folderName);

        // Crear la carpeta si no existe
        File folder = new File(folderPath);
        folder.mkdirs();


        String formattedDate = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss").format(new Date());
        String excelFileName = fileName.replace(".xlsx", "-" + folderName + "-" + formattedDate + ".xlsx");


        // Agregar "-estatus" al nombre del archivo Excel
        String excelFilePath = folderPath + File.separator + excelFileName;

        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {

            Sheet sheet = workbook.createSheet("LogSheet");

            int rowNum = 0;
            for (String message : messages) {
                Row row = sheet.createRow(rowNum++);
                Cell cell = row.createCell(0);
                cell.setCellValue(message);
            }

            workbook.write(fileOut);
            System.out.println("Mensajes registrados en el archivo Excel: " + excelFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }



}
