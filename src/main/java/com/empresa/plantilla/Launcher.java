package com.empresa.plantilla;

// Esta clase NO hereda de Application, así que Java no se asusta al verla.
// Su único trabajo es llamar a tu Main real.
public class Launcher {
    public static void main(String[] args) {
        Main.main(args);
    }
}