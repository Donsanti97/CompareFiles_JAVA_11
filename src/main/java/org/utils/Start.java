package org.utils;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.utils.MethotsAzureMasterFiles.*;

public class Start {


    public void start() {
        System.out.println("\n" +
                "  _______   ___      _________________________.____     \n" +
                " /   ___/  /  _  \\    /     \\__    _/\\_   ___/|    |    \n" +
                " \\_____  \\  /  /_\\  \\  /  \\ /  \\|    |    |    _) |    |    \n" +
                " /        \\/    |    \\/    Y    \\    |    |        \\|    |___ \n" +
                "/_______  /\\____|__  /\\____|__  /____|   /_______  /|_______ \\\n" +
                "        \\/         \\/         \\/                 \\/         \\/\n");
        System.out.println("BIENVENIDO, VAMOS A REALIZAR UN TEST DE LA DATA");
        System.out.println("Espere por favor, va iniciar el proceso");
        try {
            //Ponemos a "Dormir" el programa 5sg
            Thread.sleep(5 * 1000);
            System.out.println("Generando analisis...");
            System.console();
            excecution();
            runtime();
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public static void excecution() {
        IOUtils.setByteArrayMaxOverride(300000000);
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        try {
            JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure a analizar");
            String azureFile = getDocument();
            JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro a analizar");
            String masterFiles = getDocument();


            while (azureFile == null || masterFiles == null){
                errorMessage("No seleccionó alguno de los archivos por favor siga la instrucción a continuación");
                JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure a analizar");
                azureFile = getDocument();
                JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro a analizar");
                masterFiles = getDocument();
            }

            System.out.println("Archivos seleccionados" + "\n" +
                    new File(azureFile).getName() + "\n" + new File(masterFiles).getName());



            List<String> nameSheets1 = new ArrayList<>();
            List<String> nameSheets2 = new ArrayList<>();
            Workbook workbook = WorkbookFactory.create(new File(azureFile));
            Workbook workbook2 = WorkbookFactory.create(new File(masterFiles));
            Sheet sheet1 = null;
            Sheet sheet2 = null;

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheet1 = workbook.getSheetAt(i);
                nameSheets1.add(sheet1.getSheetName());
            }
            for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
                sheet2 = workbook2.getSheetAt(i);
                nameSheets2.add(sheet2.getSheetName());
            }

            List<String> sht1 = new ArrayList<>();
            List<String> sht2 = new ArrayList<>();
            List<String> dataList = createDualDropDownListsAndReturnSelectedValues(nameSheets1, nameSheets2);
            List<String> encabezados1;
            List<String> encabezados2;
            String encabezado = "";
            String codigo1 = "";
            String fechacorteAF = "";
            String codigo2 = "";
            String fechaCorteMF = "";
            String message;
            for (String seleccion : dataList) {
                String[] elementos = seleccion.split(SPECIAL_CHAR);

                sht1.add(elementos[0]);
                sht2.add(elementos[1]);



            }
            for (int i = 0; i < dataList.size(); i++) {
                sheet1 = workbook.getSheet(sht1.get(i));
                sheet2 = workbook2.getSheet(sht2.get(i));
            }

            if (sheet1 != null && sheet2 != null) {
                encabezados1 = getHeadersN(sheet1);
                JOptionPane.showMessageDialog(null, "Del siguiente menú escoja el primer encabezado ubicado en las hojas del archivo Maestro");
                encabezado = mostrarMenu(encabezados1);
                while (encabezado == null) {
                    errorMessage("No fue seleccionado el encabezado. Por favor siga la instrucción");
                    JOptionPane.showMessageDialog(null, "Del siguiente menú escoja el primer encabezado ubicado en las hojas del archivo Maestro");
                    encabezado = mostrarMenu(encabezados1);
                }
                encabezados2 = getHeadersMasterfile(sheet1, sheet2, encabezado);
                JOptionPane.showMessageDialog(null, "Seleccione el encabezado \"Código\" del archivo Azure que será usado para el análisis entre hojas");
                codigo1 = mostrarMenu(encabezados1);
                while (codigo1 == null) {
                    errorMessage("No fue seleccionado el código. Por favor siga la instrucción");
                    JOptionPane.showMessageDialog(null, "Seleccione el encabezado \"Código\" del archivo Azure que será usado para el análisis entre hojas");
                    codigo1 = mostrarMenu(encabezados1);
                }
                JOptionPane.showMessageDialog(null, "Seleccione el encabezado de la \"fecha de corte\" del archivo Azure que desee compara entre los dos archivos");
                fechacorteAF = mostrarMenu(encabezados1);
                if (fechacorteAF == null || fechacorteAF.equals("Nunguno")) {
                    errorMessage("No fue seleccionado la fecha de corte. Por favor siga la instrucción");
                    JOptionPane.showMessageDialog(null, "Seleccione el encabezado de la \"fecha de corte\" del archivo Azure que desee compara entre los dos archivos");
                    fechacorteAF = mostrarMenu(encabezados1);
                }
                JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda al \"Código\" del archivo Maestro que será analizado");
                codigo2 = mostrarMenu(encabezados2);
                while (codigo2 == null) {
                    errorMessage("No fue seleccionado el el código. Por favor siga la instrucción");
                    JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda al \"Código\" del archivo Maestro que será analizado");
                    codigo2 = mostrarMenu(encabezados2);
                }
                JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda a la \"Fecha de corte\" del archivo Maestro que será analizada");
                fechaCorteMF = mostrarMenu(encabezados2);
                if (fechaCorteMF == null || fechaCorteMF.equals("Nunguno")) {
                    errorMessage("No fue seleccionado la fecha de corte. Por favor siga la instrucción");
                    JOptionPane.showMessageDialog(null, "Seleccione el encabezado que corresponda a la \"Fecha de corte\" del archivo Maestro que será analizada");
                    fechaCorteMF = mostrarMenu(encabezados2);
                }
            }

            String fecha = parsearFecha(fechaCorteMF);
            System.out.println("Fecha modificada: " + fecha);


            for (int i = 0; i < dataList.size(); i++) {
                sheet1 = workbook.getSheet(sht1.get(i));
                sheet2 = workbook2.getSheet(sht2.get(i));

                if ((fechacorteAF == null || fechacorteAF.equals("Nunguno")) ||
                        (fechaCorteMF == null || fechaCorteMF.equals("Nunguno"))){
                    errorMessage("Las hojas " + sheet1.getSheetName() + " - " + sheet2.getSheetName() + ", NO pueden ser analizadas, la información es incompleta." +
                            "\n Por favor verifique que los archivos tengan la información necesaria para el análisis");
                } else if (!fechacorteAF.equals(fecha)) {
                    String yesNoAnswer = showYesNoDialog("Los dos valores que intenta comparar estan contenidos en un encabezado tipo fecha de corte " +
                            "\n O diferente al que seleccionó al comienzo del programa? ?");
                    if (yesNoAnswer.equals("SI")){
                        encabezados1 = getHeadersN(sheet1);
                        encabezados2 = getHeadersMasterfile(sheet1, sheet2, encabezado);

                        JOptionPane.showMessageDialog(null, "Seleccione el encabezado del archivo Azure que desee compara entre los dos archivos");
                        fechacorteAF = mostrarMenu(encabezados1);
                        if (fechacorteAF == null || fechacorteAF.equals("Nunguno")) {
                            errorMessage("No fue seleccionado la fecha de corte. Por favor siga la instrucción");
                            JOptionPane.showMessageDialog(null, "Seleccione el encabezado del archivo Azure que desee compara entre los dos archivos");
                            fechacorteAF = mostrarMenu(encabezados1);
                        }

                        JOptionPane.showMessageDialog(null, "Seleccione el encabezado del archivo Maestro que será analizada");
                        fechaCorteMF = mostrarMenu(encabezados2);
                        if (fechaCorteMF == null || fechaCorteMF.equals("Nunguno")) {
                            errorMessage("No fue seleccionado la fecha de corte. Por favor siga la instrucción");
                            JOptionPane.showMessageDialog(null, "Seleccione el encabezado del archivo Maestro que será analizada");
                            fechaCorteMF = mostrarMenu(encabezados2);
                        }
                        messageSheets(sheet1, sheet2, workbook, workbook2, codigo1, fechacorteAF, codigo2, fechaCorteMF, encabezado);

                    } else {
                        message = "La fecha de corte que intenta validar no se encuentra. \n La hoja [" + sheet2.getSheetName() + "] no se podrá analizar";
                        errorMessage(message);
                        System.out.println(message);
                    }

                }else {
                    messageSheets(sheet1, sheet2, workbook, workbook2, codigo1, fechacorteAF, codigo2, fechaCorteMF, encabezado);

                }
            }
            System.out.println("---------------------------------------------------------------------------------------");
            System.out.println("Analisis completado...");
            logWinsToFile(azureFile, coincidencias);
            logErrorsToFile(azureFile, errores);
            workbook.close();
            workbook2.close();


            //moveDocument(file2, destino);

            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente sin errores");


        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

    public static void messageSheets(Sheet sheet1, Sheet sheet2,
                                       Workbook workbook1, Workbook workbook2,
                                     String codigo1, String fechacorteAF,
                                     String codigo2, String fechaCorteMF, String encabezado){
        String message;
        List<Map<String, String>> valoresEncabezados1;
        List<Map<String, String>> valoresEncabezados2;
        List<Map<String, String>> mapList1;
        List<Map<String, String>> mapList2;
        List<String> encabezados1 = getHeadersN(sheet1);

        try {

            /*------------------------------------------------------------------------------------------------*/
            valoresEncabezados1 = obtenerValoresPorFilas(sheet1, encabezados1, codigo1, fechacorteAF);
            valoresEncabezados2 = obtenerValoresPorFilas(workbook1, workbook2, sheet1.getSheetName(), sheet2.getSheetName(), codigo2, fechaCorteMF, encabezado);

            if (valoresEncabezados1 != null && valoresEncabezados2 != null) {
                mapList1 = createMapList(valoresEncabezados1, codigo1, fechacorteAF);
                mapList2 = createMapList(valoresEncabezados2, codigo2, fechaCorteMF);
                for (Map<String, String> map1 : mapList1) {
                    for (Map.Entry<String, String> entry1 : map1.entrySet()) {
                        for (Map<String, String> map2 : mapList2) {
                            for (Map.Entry<String, String> entry2 : map2.entrySet()) {


                                if (entry1.getKey().equalsIgnoreCase(entry2.getKey())) {

                                    System.out.println("CÓDIGO ENCONTRADO: " + entry1.getKey());

                                    String mapList2Value;
                                    String mapList1Value;

                                    if (entry2.getValue().contains(entry1.getValue())){
                                        message = sheet1.getSheetName() + " - " + sheet2.getSheetName() +
                                                "\n" + entry1.getKey() + "-> Los valores: " + entry1.getValue() + " & " + entry2.getValue() + " coinciden";
                                        coincidencias.add(message);
                                        System.out.println(message);
                                    } else {
                                        mapList2Value = ponerDecimales(entry2.getValue());
                                        mapList1Value = ponerDecimales(entry1.getValue());
                                        if (mapList1Value.equals(mapList2Value)) {
                                            message = sheet1.getSheetName() + " - " + sheet2.getSheetName() +
                                                    "\n" + entry1.getKey() + "-> Los valores: " + entry1.getValue() + " & " + entry2.getValue() + " coinciden";
                                            coincidencias.add(message);
                                            System.out.println(message);
                                        } else if (mapList2Value.contains(mapList1Value)) {
                                            message = sheet1.getSheetName() + " - " + sheet2.getSheetName() +
                                                    "\n" + entry1.getKey() + "-> Los valores: " + entry1.getValue() + " & " + entry2.getValue() + " puede que sean iguales";
                                            coincidencias.add(message);
                                            System.out.println(message);
                                        } else {
                                            message = sheet1.getSheetName() + " - " + sheet2.getSheetName() +
                                                    "\n" + entry1.getKey() + "-> Los valores: " + entry1.getValue() + " & " + entry2.getValue() + " NO coinciden";
                                            errores.add(message);
                                            System.out.println(message);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            } else {
                errorMessage("No es posible analizar los valores ya que los campos están incompletos." +
                        "\n Por favor verifique la información de los archivo con respecto a las hojas " +
                        sheet1.getSheetName() + ", " + sheet2.getSheetName());
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    public static String parsearFecha(String fechaString) {
        SimpleDateFormat formatoEntrada = new SimpleDateFormat("dd-MMM-yy", new Locale("es", "ES"));
        SimpleDateFormat formatoSalida = new SimpleDateFormat("dd/MM/yyyy");

        try {
            if (fechaString == null || fechaString.equals("Ninguno")) {
                System.err.println("Fecha no encontrada");
                return null;
            } else {
                Date fecha = formatoEntrada.parse(fechaString);
                return formatoSalida.format(fecha);
            }
        } catch (ParseException e) {
            e.printStackTrace();
        }

        return null; // o manejar el error de alguna manera
    }

}
