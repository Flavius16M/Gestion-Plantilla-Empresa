package com.empresa.plantilla;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.util.Scanner;

// Clase principal del programa. Muestra el menú por consola y gestiona la interacción con el usuario.
public class Main {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in, "UTF-8");
        GestorPlantilla gestor = new GestorPlantilla();
        ConfigManager config = new ConfigManager();

        System.out.println("╔══════════════════════════════════════════════╗");
        System.out.println("║     GESTIÓN DE PLANTILLA - ROTACIÓN PERSONAL ║");
        System.out.println("╚══════════════════════════════════════════════╝");
        System.out.println();

        // Si es la primera vez que se abre el programa, pedimos el archivo Excel
        if (!config.tieneRutaValida()) {
            System.out.println("Primera ejecución: selecciona el archivo Excel.");
            System.out.println();
            configurarRuta(scanner, config);
        } else {
            System.out.println("  Excel configurado: " + config.getRutaExcel());
            System.out.println();
        }

        boolean ejecutando = true;

        while (ejecutando) {
            System.out.println("┌─────────────────────────────────────────────┐");
            System.out.println("│  ¿Qué deseas hacer?                         │");
            System.out.println("│  1. Registrar ENTRADA de trabajador          │");
            System.out.println("│  2. Registrar SALIDA de trabajador           │");
            System.out.println("│  3. Ver plantilla actual                     │");
            System.out.println("│  4. Cambiar archivo Excel                    │");
            System.out.println("│  0. Salir                                    │");
            System.out.println("└─────────────────────────────────────────────┘");
            System.out.print("Opción: ");

            String opcion = scanner.nextLine().trim();

            switch (opcion) {
                case "1":
                    registrarEntrada(scanner, gestor, config.getRutaExcel());
                    break;
                case "2":
                    registrarSalida(scanner, gestor, config.getRutaExcel());
                    break;
                case "3":
                    try {
                        gestor.cargarExcel(config.getRutaExcel());
                        gestor.mostrarPlantillaActual();
                    } catch (Exception e) {
                        System.out.println("[ERROR] " + e.getMessage() + "\n");
                    }
                    break;
                case "4":
                    configurarRuta(scanner, config);
                    break;
                case "0":
                    System.out.println("\nCerrando programa. ¡Hasta luego!");
                    ejecutando = false;
                    break;
                default:
                    System.out.println("\n[!] Opción no válida.\n");
            }
        }

        scanner.close();
    }

    // Abre un selector de archivos gráfico para elegir el Excel.
    // Usamos una ventana de Windows en vez de escribir la ruta a mano para evitar
    // problemas con nombres de archivo que tienen tildes o caracteres especiales.
    private static void configurarRuta(Scanner scanner, ConfigManager config) {
        System.out.println("\n── SELECCIONAR ARCHIVO EXCEL ──────────────────");
        System.out.println("Se abrirá una ventana para seleccionar el archivo.");
        System.out.println("Pulsa Enter para abrir el selector...");
        scanner.nextLine();

        File archivo = abrirSelectorArchivo();

        if (archivo == null) {
            System.out.println("[!] No se seleccionó ningún archivo.\n");
            return;
        }

        config.setRutaExcel(archivo.getAbsolutePath());
        System.out.println("✔ Archivo seleccionado : " + archivo.getName());
        System.out.println("✔ Ruta guardada. No volverá a pedirse.\n");
    }

    private static File abrirSelectorArchivo() {
        // Usamos el look and feel del sistema para que la ventana tenga el aspecto de Windows
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ignored) {}

        JFileChooser selector = new JFileChooser();
        selector.setDialogTitle("Selecciona el archivo Excel de rotación de personal");
        selector.setFileSelectionMode(JFileChooser.FILES_ONLY);
        selector.setFileFilter(new FileNameExtensionFilter("Archivos Excel (*.xlsx)", "xlsx"));

        // Intentamos abrir el selector directamente en la carpeta Documentos
        File documentos = new File(System.getProperty("user.home"), "Documents");
        if (!documentos.exists()) documentos = new File(System.getProperty("user.home"), "Documentos");
        if (documentos.exists()) selector.setCurrentDirectory(documentos);

        int resultado = selector.showOpenDialog(null);
        if (resultado == JFileChooser.APPROVE_OPTION) {
            return selector.getSelectedFile();
        }
        return null;
    }

    private static void registrarEntrada(Scanner scanner, GestorPlantilla gestor, String ruta) {
        System.out.println("\n── NUEVA ENTRADA ──────────────────────────────");
        System.out.print("Nombre del trabajador que ENTRA : ");
        String iniciales = scanner.nextLine().trim().toUpperCase();

        if (iniciales.isEmpty()) {
            System.out.println("[!] El nombre no puede estar vacío.\n");
            return;
        }

        System.out.print("Fecha del cambio (dd/MM/yyyy) [Enter = hoy]: ");
        String fechaStr = scanner.nextLine().trim();

        try {
            gestor.registrarEntrada(ruta, iniciales, fechaStr);
        } catch (Exception e) {
            System.out.println("[ERROR] " + e.getMessage());
        }
        System.out.println();
    }

    private static void registrarSalida(Scanner scanner, GestorPlantilla gestor, String ruta) {
        System.out.println("\n── NUEVA SALIDA ───────────────────────────────");

        // Cargamos y mostramos la plantilla actual antes de pedir las iniciales,
        // para que el usuario pueda ver qué trabajadores hay en ese momento
        try {
            gestor.cargarExcel(ruta);
            gestor.mostrarPlantillaActual();
        } catch (Exception e) {
            System.out.println("[ERROR] No se pudo cargar el Excel: " + e.getMessage() + "\n");
            return;
        }

        System.out.print("Nombre del trabajador que SALE: ");
        String iniciales = scanner.nextLine().trim().toUpperCase();

        if (iniciales.isEmpty()) {
            System.out.println("[!] El nombre no puede estar vacío.\n");
            return;
        }

        System.out.print("Fecha del cambio (dd/MM/yyyy) [Enter = hoy]: ");
        String fechaStr = scanner.nextLine().trim();

        try {
            gestor.registrarSalida(ruta, iniciales, fechaStr);
        } catch (Exception e) {
            System.out.println("[ERROR] " + e.getMessage());
        }
        System.out.println();
    }
}
