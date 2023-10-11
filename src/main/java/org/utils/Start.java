package org.utils;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.List;
import java.util.Map;

import static org.utils.MethotsAzureMasterFiles.*;

public class Start {
    public void start() {
        //MethotsAzureMasterFiles readFiles = new MethotsAzureMasterFiles();
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

    public static void moveDocument(String origen, String destino) throws IOException {
        Path origenPath = Paths.get(origen);
        Path destinoPath = Paths.get(destino);

        // Mueve el documento desde la ubicación de origen a la ubicación de destino
        Files.move(origenPath, destinoPath, StandardCopyOption.REPLACE_EXISTING);
    }

    public static void excecution() {
        try {
            JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure a analizar");
            String file1 = getDocument();
            JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro a analizar");
            String file2 = getDocument();
            File file = new File(file2);
            String destino = System.getProperty("user.home") + File.separator + "Documentos" + File.separator + "procesedDocuments" + File.separator + file.getName();


            if (file1 != null && file2 != null) {
                System.out.println("Ruta del archivo Excel seleccionado: " + file1);
                System.out.println("Ruta del archivo Excel seleccionado: " + file2);
            } else {
                System.out.println("No se seleccionó ningún archivo.");
            }


            FileInputStream fis = new FileInputStream(file1);
            Workbook workbook = new XSSFWorkbook(fis);
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook workbook2 = new XSSFWorkbook(fis2);
            Sheet sheet1 = workbook.getSheetAt(0);

            int indexF2 = 0;
            List<String> nameSheets1 = getWorkSheet(file1, 0);
            Sheet sheet2 = workbook2.getSheetAt(0);
            List<String> nameSheets2 = getWorkSheet(file2, 0);

            for (String s2 : nameSheets2) {
                String sheetname = s2.replaceAll("\\s", "");
                for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
                    if (nameSheets1.get(0).equals(sheetname)) {
                        indexF2 = i;
                        System.out.println("La hoja de trabajo se encontró en Excel B en el índice: " + indexF2);
                        break;
                    } else {
                        System.out.println("Analizando datos...");
                        break;
                    }
                }
            }

            sheet2 = workbook2.getSheetAt(indexF2);
            nameSheets2 = getWorkSheet(file2, indexF2);

            List<String> encabezados1 = null;
            List<String> encabezados2 = null;

            List<Map<String, String>> valoresEncabezados1 = null;
            List<Map<String, String>> valoresEncabezados2 = null;

            System.out.println("Analizando archivo Azure");
            for (String sheets : nameSheets1) {
                System.out.print("Analizando: ");
                System.out.println(sheets);
                encabezados1 = getHeaders(file1, sheets);
                //System.out.println("Headers: ");
                for (String headers : encabezados1) {
                    //System.out.print(headers + "||");
                    valoresEncabezados1 = getValuebyHeader(file1, sheets);
                }
            }
            System.out.println("------------------------------------------------------------------------------------------");

            valoresEncabezados1 = obtenerValoresPorFilas(sheet1, encabezados1);
            for (Map<String, String> map : valoresEncabezados1) {
                System.out.println("Analizando valores... ");
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    System.out.println("Headers1: " + entry.getKey() + ", value: " + entry.getValue());
                }
            }

            System.out.println("------------------------------------------------------");
            System.out.println("Analizando archivo Maestro");
            for (String sheets2 : nameSheets2) {
                System.out.println("Analizando: " + sheets2);
                encabezados2 = getHeadersMasterfile(sheet1, sheet2);
                for (String headers : encabezados2) {
                    System.out.println("Headers2: " + headers);
                }
            }

            System.out.println("-------------------------------------------------------------------------------------");
            valoresEncabezados2 = obtenerValoresPorFilas(sheet2, encabezados2);
            for (Map<String, String> map : valoresEncabezados2) {
                System.out.println("Analizando valores... ");
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    System.out.println("Headers2: " + entry.getKey() + ", Value: " + entry.getValue());
                }
            }

            System.out.println("---------------------------------------------------------------------------------------");
            /*for (String e1 : encabezados1) {
                for (String e2 : encabezados2) {
                    if (e1.equals(e2)) {
                        System.out.println("equals" + e1 + ", " + e2);
                    } else {
                        System.out.println("No equals");
                    }
                }
            }*/
            System.out.println("Analisis completado...");
            workbook.close();
            workbook2.close();
            fis.close();
            fis2.close();


            //moveDocument(file2, destino);

            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente sin errores");


        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
