package com.empresa.plantilla;

import atlantafx.base.theme.PrimerLight;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

// 1. Para que sea una app gráfica, Main debe heredar de "Application"
public class Main extends Application {

    // 2. Este método es el que "dibuja" la ventana al arrancar
    @Override
    public void start(Stage ventana) throws Exception {
        
        //Esto le da el aspecto de app moderna (estilo Windows 11 / GitHub)
        Application.setUserAgentStylesheet(new PrimerLight().getUserAgentStylesheet());

        // Buscamos el archivo que guardaste en la carpeta resources
        FXMLLoader loader = new FXMLLoader(getClass().getResource("/pantalla_principal.fxml"));
        
        // Metemos tu diseño en una "Escena" (le damos un tamaño por defecto de 400x350 píxeles)
        Scene escena = new Scene(loader.load(), 400, 350);

        // Configuramos la ventana
        ventana.setTitle("Gestión de Plantilla");
        ventana.setScene(escena);
        ventana.show(); // ¡Que se haga la luz!
    }

    // 3. El main tradicional ahora solo lanza la app gráfica
    public static void main(String[] args) {
        launch(args); // Esto arranca la aplicación gráfica
    }
}