package com.empresa.plantilla;

import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.TextInputDialog;
import javafx.stage.FileChooser;
import java.io.File;
import java.util.Optional;

public class MenuController {

    private GestorPlantilla gestor = new GestorPlantilla();
    private ConfigManager config = new ConfigManager();

    // ----------------------------------------------------------------
    // BOTÓN 1: REGISTRAR ENTRADA
    // ----------------------------------------------------------------
    @FXML
    public void btnEntradaClick() {
        if (!config.tieneRutaValida()) {
            mostrarError("Primero debes seleccionar el archivo Excel con el botón de abajo.");
            return;
        }

        // 1. Pedir Iniciales (Sustituye a tu antiguo Scanner)
        TextInputDialog dialogNombre = new TextInputDialog();
        dialogNombre.setTitle("Registrar Entrada");
        dialogNombre.setHeaderText("Alta de nuevo trabajador");
        dialogNombre.setContentText("Iniciales del trabajador:");
        Optional<String> resultNombre = dialogNombre.showAndWait();

        if (resultNombre.isPresent() && !resultNombre.get().trim().isEmpty()) {
            String iniciales = resultNombre.get().trim().toUpperCase();

            // 2. Pedir Fecha
            TextInputDialog dialogFecha = new TextInputDialog();
            dialogFecha.setTitle("Registrar Entrada");
            dialogFecha.setHeaderText("Fecha de incorporación");
            dialogFecha.setContentText("Fecha (dd/MM/yyyy) [Deja en blanco para HOY]:");
            Optional<String> resultFecha = dialogFecha.showAndWait();

            if (resultFecha.isPresent()) {
                try {
                    // 3. ¡AQUÍ ESTÁ LA MAGIA! Llamamos a tu código del Excel de siempre
                    gestor.registrarEntrada(config.getRutaExcel(), iniciales, resultFecha.get().trim());
                    mostrarExito("¡El trabajador " + iniciales + " se ha añadido al Excel correctamente!");
                } catch (Exception e) {
                    mostrarError("Error al modificar el Excel: " + e.getMessage());
                }
            }
        }
    }

    // ----------------------------------------------------------------
    // BOTÓN 2: REGISTRAR SALIDA
    // ----------------------------------------------------------------
    @FXML
    public void btnSalidaClick() {
        if (!config.tieneRutaValida()) {
            mostrarError("Primero debes seleccionar el archivo Excel.");
            return;
        }

        TextInputDialog dialogNombre = new TextInputDialog();
        dialogNombre.setTitle("Registrar Salida");
        dialogNombre.setHeaderText("Baja de trabajador");
        dialogNombre.setContentText("Iniciales del trabajador que sale:");
        Optional<String> resultNombre = dialogNombre.showAndWait();

        if (resultNombre.isPresent() && !resultNombre.get().trim().isEmpty()) {
            String iniciales = resultNombre.get().trim().toUpperCase();

            TextInputDialog dialogFecha = new TextInputDialog();
            dialogFecha.setTitle("Registrar Salida");
            dialogFecha.setHeaderText("Fecha de salida");
            dialogFecha.setContentText("Fecha (dd/MM/yyyy) [Deja en blanco para HOY]:");
            Optional<String> resultFecha = dialogFecha.showAndWait();

            if (resultFecha.isPresent()) {
                try {
                    gestor.registrarSalida(config.getRutaExcel(), iniciales, resultFecha.get().trim());
                    mostrarExito("¡El trabajador " + iniciales + " ha sido dado de baja en el Excel!");
                } catch (Exception e) {
                    mostrarError("Error al modificar el Excel: " + e.getMessage());
                }
            }
        }
    }

    // ----------------------------------------------------------------
    // BOTÓN 3: VER PLANTILLA (Este te lo dejo de deberes)
    // ----------------------------------------------------------------
   @FXML
public void btnVerPlantillaClick() {
    String ruta = config.getRutaExcel();
    if (ruta == null || ruta.isEmpty()) {
        mostrarError("Selecciona primero el archivo Excel.");
        return;
    }
    
    String resultado = gestor.getPlantillaComoTexto(ruta);
    
    // ... aquí el código de la ventana con el TextArea que pusimos antes ...
    mostrarVentanaTexto("Plantilla Actual", resultado);
}
    

    // ----------------------------------------------------------------
    // BOTÓN 4: CAMBIAR EXCEL (Tu antiguo JFileChooser)
    // ----------------------------------------------------------------
    @FXML
    public void btnCambiarExcelClick() {
        FileChooser selector = new FileChooser();
        selector.setTitle("Seleccionar Excel de la Plantilla");
        // Hacemos que solo le deje elegir archivos Excel
        selector.getExtensionFilters().add(new FileChooser.ExtensionFilter("Archivos Excel", "*.xlsx"));
        
        File archivoElegido = selector.showOpenDialog(null);
        
        if (archivoElegido != null) {
            config.setRutaExcel(archivoElegido.getAbsolutePath());
            mostrarExito("¡Archivo vinculado con éxito!\n" + archivoElegido.getName());
        }
    }

    // ================================================================
    // MÉTODOS DE AYUDA PARA MOSTRAR MENSAJES BONITOS
    // ================================================================
    private void mostrarExito(String mensaje) {
        Alert alerta = new Alert(Alert.AlertType.INFORMATION);
        alerta.setTitle("Éxito");
        alerta.setHeaderText(null);
        alerta.setContentText(mensaje);
        alerta.showAndWait();
    }

    private void mostrarError(String mensaje) {
        Alert alerta = new Alert(Alert.AlertType.ERROR);
        alerta.setTitle("Error");
        alerta.setHeaderText("Ha ocurrido un problema");
        alerta.setContentText(mensaje);
        alerta.showAndWait();
    }
    // Este método es el que crea la ventana con el texto ordenado
private void mostrarVentanaTexto(String titulo, String contenido) {
    Alert alert = new Alert(Alert.AlertType.INFORMATION);
    alert.setTitle(titulo);
    alert.setHeaderText(null);

    javafx.scene.control.TextArea textArea = new javafx.scene.control.TextArea(contenido);
    textArea.setEditable(false);
    textArea.setPrefHeight(400);
    textArea.setPrefWidth(450);
    
    // IMPORTANTE: Fuente Consolas para que los nombres salgan en columna perfecta
    textArea.setStyle("-fx-font-family: 'Consolas'; -fx-font-size: 13px;");

    alert.getDialogPane().setContent(textArea);
    alert.showAndWait();
}
}