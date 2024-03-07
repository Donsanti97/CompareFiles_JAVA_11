package org.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Font;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.*;

public class MethotsAzureMasterFiles {

    public static final String SPECIAL_CHAR = " -X- ";
    public static List<String> errores = new ArrayList<>();
    public  static List<String> coincidencias = new ArrayList<>();

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
            return archivoSeleccionado.getAbsolutePath();
        } else {
            return null; // Si no se seleccionó ningún archivo, retorna null
        }
    }

    /*-------------------------------------------------------------------------------------------------------------------------------*/
    public static void runtime() {
        Runtime runtime = Runtime.getRuntime();
        long minRunningMemory = (8L * 1024L * 1024L * 1024L);
        if (runtime.freeMemory() < minRunningMemory) {
            System.gc();
        }
    }
    /*---------------------------------------------------------------------------------------------------------------*/

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

    public static String obtenerValorVisibleCelda(Cell cell) {
        try {
            DataFormatter dataFormatter = new DataFormatter();
            String valor;

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
                        //String dataFormatString = cell.getCellStyle().getDataFormatString();

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
                    valor = "0";
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
                    if (key.equals("null")) {
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

    public static void errorMessage(String mensaje) {
        JLabel label = new JLabel("<html><font color='red'>" + mensaje + "</font></html>");
        label.setFont(new Font("Arial", Font.PLAIN, 14)); // Puedes ajustar la fuente según tus preferencias

        JOptionPane.showMessageDialog(null, label, "Error", JOptionPane.ERROR_MESSAGE);
    }


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
