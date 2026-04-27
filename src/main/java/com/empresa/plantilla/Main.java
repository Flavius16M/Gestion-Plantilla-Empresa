package com.empresa.plantilla;

import atlantafx.base.theme.PrimerLight;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class Main extends Application {

    @Override
    public void start(Stage ventana) throws Exception {

        Application.setUserAgentStylesheet(
            new PrimerLight().getUserAgentStylesheet()
        );

        FXMLLoader loader = new FXMLLoader(
            getClass().getResource("/pantalla_principal.fxml")
        );

        Parent root = loader.load();

        Scene escena = new Scene(root);

        ventana.setTitle("Gestión de Plantilla");
        ventana.setScene(escena);

        ventana.setWidth(820);
        ventana.setHeight(320);

        ventana.setResizable(false);

        ventana.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}