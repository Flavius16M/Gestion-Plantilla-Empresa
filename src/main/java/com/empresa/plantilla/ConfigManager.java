package com.empresa.plantilla;

import java.io.*;
import java.util.Properties;

// Esta clase se encarga de guardar y leer la configuración del programa.
// La configuración es básicamente la ruta del archivo Excel que usa el programa.
// Se guarda en un archivo .properties para que no haya que volver a introducirla cada vez.
public class ConfigManager {

    private static final String ARCHIVO_CONFIG = "config.properties";
    private static final String CLAVE_RUTA     = "ruta.excel";

    private final Properties props = new Properties();
    private final File archivoConfig;

    public ConfigManager() {
        // Buscamos una carpeta donde podamos escribir sin problemas de permisos
        File dir = encontrarCarpetaEscribible();
        archivoConfig = new File(dir, ARCHIVO_CONFIG);
        System.out.println("  [Config] " + archivoConfig.getAbsolutePath());
        cargar();
    }

    // Prueba varias carpetas del sistema y devuelve la primera donde se puede escribir.
    // Esto es necesario porque en entornos corporativos no siempre se puede escribir en cualquier sitio.
    private File encontrarCarpetaEscribible() {
        String home = System.getProperty("user.home");
        String tmp  = System.getProperty("java.io.tmpdir");

        String[] candidatos = {
            home + File.separator + "Documents"  + File.separator + "GestionPlantilla",
            home + File.separator + "Documentos" + File.separator + "GestionPlantilla",
            home + File.separator + "GestionPlantilla",
            tmp  + File.separator + "GestionPlantilla",
        };

        for (String ruta : candidatos) {
            File dir = new File(ruta);
            dir.mkdirs();
            if (puedeEscribir(dir)) {
                return dir;
            }
        }

        // Si ninguna funciona, usamos la carpeta temporal del sistema como último recurso
        File fallback = new File(tmp, "GestionPlantilla");
        fallback.mkdirs();
        return fallback;
    }

    // Crea un archivo de prueba para comprobar si tenemos permisos de escritura en esa carpeta
    private boolean puedeEscribir(File dir) {
        File test = new File(dir, ".test_escritura");
        try {
            if (test.createNewFile()) {
                test.delete();
                return true;
            }
        } catch (IOException e) {
            // Si lanza excepción es que no tenemos permisos
        }
        return false;
    }

    public String getRutaExcel() {
        return props.getProperty(CLAVE_RUTA);
    }

    public void setRutaExcel(String ruta) {
        props.setProperty(CLAVE_RUTA, ruta);
        guardar();
    }

    // Comprueba que la ruta guardada no esté vacía y que el archivo realmente exista
    public boolean tieneRutaValida() {
        String ruta = getRutaExcel();
        if (ruta == null || ruta.isEmpty()) return false;
        return new File(ruta).exists();
    }

    private void cargar() {
        if (archivoConfig.exists()) {
            try (FileInputStream fis = new FileInputStream(archivoConfig)) {
                props.load(fis);
            } catch (IOException e) {
                // Si no se puede leer, el programa arranca sin configuración guardada
            }
        }
    }

    private void guardar() {
        try (FileOutputStream fos = new FileOutputStream(archivoConfig)) {
            props.store(fos, "Configuracion Gestion Plantilla");
        } catch (IOException e) {
            System.out.println("[ERROR al guardar config] " + e.getMessage());
        }
    }
}
