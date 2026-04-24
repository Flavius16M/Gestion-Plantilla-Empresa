package com.empresa.plantilla;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

// Clase principal que maneja toda la lógica con el archivo Excel.
// Usa la librería Apache POI para leer y escribir archivos .xlsx.
public class GestorPlantilla {

    // Posición de cada columna en el Excel (índice 0 = columna A, 1 = B, etc.)
    static final int COL_AÑO      = 1;  // B - año (ej: AÑO 2026)
    static final int COL_TRIMESTRE = 2;  // C - trimestre (ej: 2T 2026)
    static final int COL_PCT_TRIM  = 3;  // D - porcentaje trimestral
    static final int COL_FECHA     = 4;  // E - fecha del cambio
    static final int COL_NUM_PERS  = 5;  // F - número de personas en plantilla
    static final int COL_INICIALES = 6;  // G en adelante - iniciales de cada trabajador

    static final int FILA_NUMEROS   = 2; // fila 3: numeración de posiciones (1, 2, 3...)
    static final int FILA_CABECERA  = 3; // fila 4: cabeceras de las columnas
    static final int FILA_DATOS_INI = 4; // fila 5: aquí empiezan los datos reales

    // Colores que se usan en el Excel
    private static final byte[] COLOR_GRIS    = hexToRgb("d8d8d8"); // huecos sin trabajador
    private static final byte[] COLOR_VERDE   = hexToRgb("92d050"); // máximo del trimestre
    private static final byte[] COLOR_NARANJA = hexToRgb("ffc000"); // mínimo del trimestre

    private static final SimpleDateFormat FMT = new SimpleDateFormat("dd/MM/yyyy");

    private Workbook workbook;
    private Sheet hoja;
    private String rutaCargada;
    

    // Estilos del Excel definidos aquí arriba para reutilizarlos en cada fila que se escribe.
    // Es importante no crearlos dentro de un bucle porque Excel tiene un límite de estilos
    // y el programa daría error si se superara ese límite.
    private CellStyle estiloAÑO;
    private CellStyle estiloTrimestre;
    private CellStyle estiloTrimestreUnico; // para cuando un trimestre solo tiene una fila
    private CellStyle estiloPct;
    private CellStyle estiloFecha;
    private CellStyle estiloNumero;
    private CellStyle estiloInicial;
    private CellStyle estiloGris;
    private CellStyle estiloVerde;
    private CellStyle estiloNaranja;
    private List<String> plantillaActual = new ArrayList<>();
    

    // Clase interna que representa una fila del Excel:
    // guarda la fecha del cambio y la lista completa de trabajadores en ese momento.
    private static class RegistroFila {
        final Date fecha;
        final List<String> personal;

        RegistroFila(Date fecha, List<String> personal) {
            this.fecha = fecha;
            this.personal = new ArrayList<>(personal);
        }
    }

    // ── Métodos públicos ─────────────────────────────────────────────────────

    public void cargarExcel(String ruta) throws IOException {
        File f = new File(ruta);
        if (!f.exists()) throw new IOException("El archivo no existe: " + ruta);
        try (FileInputStream fis = new FileInputStream(f)) {
            workbook = new XSSFWorkbook(fis);
        }
        hoja = workbook.getSheetAt(0);
        rutaCargada = ruta;
    }

    public void registrarEntrada(String ruta, String iniciales, String fechaStr) throws Exception {
        cargarExcel(ruta);
        inicializarEstilos();
        Date fecha = parsearFecha(fechaStr);

        List<RegistroFila> registros = leerTodosLosRegistros();
        List<String> plantillaBase = getPlantillaEnFecha(registros, fecha);
        plantillaBase.add(iniciales);

        registros.add(new RegistroFila(fecha, plantillaBase));
        registros.sort(Comparator.comparing(r -> r.fecha));

        // Propagación hacia el futuro: si la fecha introducida es anterior a registros
        // que ya existen, hay que añadir al trabajador en todos los registros posteriores
        // donde todavía no aparezca (porque en ese momento ya formaba parte de la plantilla).
        boolean encontrado = false;
        for (RegistroFila reg : registros) {
            if (!encontrado && reg.fecha.equals(fecha) && reg.personal.contains(iniciales)) {
                encontrado = true;
                continue;
            }
            if (encontrado && !reg.personal.contains(iniciales)) {
                reg.personal.add(iniciales);
            }
        }

        reescribirExcel(registros);
        guardar(ruta);

        System.out.println("\n✔ Entrada registrada correctamente.");
        System.out.println("  Trabajador  : " + iniciales);
        System.out.println("  Plantilla   : " + plantillaBase.size() + " personas");
        System.out.println("  Trimestre   : " + getTrimestre(fecha) + " " + getAÑO(fecha));
        System.out.println("  Fecha       : " + FMT.format(fecha));
    }

    public void registrarSalida(String ruta, String iniciales, String fechaStr) throws Exception {
        if (workbook == null || !ruta.equals(rutaCargada)) cargarExcel(ruta);
        inicializarEstilos();
        Date fecha = parsearFecha(fechaStr);

        List<RegistroFila> registros = leerTodosLosRegistros();
        List<String> plantillaBase = getPlantillaEnFecha(registros, fecha);

        int idx = -1;
        for (int i = 0; i < plantillaBase.size(); i++) {
            if (plantillaBase.get(i).equalsIgnoreCase(iniciales)) { idx = i; break; }
        }
        if (idx == -1) {
            System.out.println("\n[!] No se encontró '" + iniciales + "' en la plantilla de esa fecha.");
            return;
        }
        plantillaBase.remove(idx);

        registros.add(new RegistroFila(fecha, plantillaBase));
        registros.sort(Comparator.comparing(r -> r.fecha));

        // Igual que en la entrada pero al revés: si la baja fue antes de registros ya existentes,
        // hay que quitar al trabajador de todos los registros posteriores donde aún aparezca.
        boolean encontrado = false;
        for (RegistroFila reg : registros) {
            if (!encontrado && reg.fecha.equals(fecha) && !reg.personal.contains(iniciales)) {
                encontrado = true;
                continue;
            }
            if (encontrado) {
                reg.personal.removeIf(p -> p.equalsIgnoreCase(iniciales));
            }
        }

        reescribirExcel(registros);
        guardar(ruta);

        System.out.println("\n✔ Salida registrada correctamente.");
        System.out.println("  Trabajador  : " + iniciales + " (eliminado)");
        System.out.println("  Plantilla   : " + plantillaBase.size() + " personas");
        System.out.println("  Trimestre   : " + getTrimestre(fecha) + " " + getAÑO(fecha));
        System.out.println("  Fecha       : " + FMT.format(fecha));
    }

    public void mostrarPlantillaActual() {
        if (hoja == null) { System.out.println("[!] Primero carga un Excel."); return; }
        List<RegistroFila> registros = leerTodosLosRegistros();
        if (registros.isEmpty()) { System.out.println("[!] No hay datos en el Excel."); return; }
        RegistroFila ultimo = registros.get(registros.size() - 1);

        System.out.println("\n── PLANTILLA ACTUAL ────────────────────────────");
        System.out.println("  Última actualización : " + FMT.format(ultimo.fecha));
        System.out.println("  Trimestre            : " + getTrimestre(ultimo.fecha) + " " + getAÑO(ultimo.fecha));
        System.out.println("  Nº personas          : " + ultimo.personal.size());
        System.out.print("  Trabajadores         : ");
        for (int i = 0; i < ultimo.personal.size(); i++) {
            System.out.print(ultimo.personal.get(i));
            if (i < ultimo.personal.size() - 1) System.out.print(", ");
        }
        System.out.println("\n");
    }

    // ── Lectura del Excel ────────────────────────────────────────────────────

    // Lee todas las filas de datos del Excel y las devuelve ordenadas por fecha.
    // Salta las filas vacías o que no tengan fecha válida.
    private List<RegistroFila> leerTodosLosRegistros() {
        List<RegistroFila> registros = new ArrayList<>();

        for (int r = FILA_DATOS_INI; r <= hoja.getLastRowNum(); r++) {
            Row row = hoja.getRow(r);
            if (row == null) continue;
            Cell cFecha = row.getCell(COL_FECHA);
            if (cFecha == null || cFecha.getCellType() == CellType.BLANK) continue;
            if (cFecha.getCellType() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(cFecha)) continue;

            Date fecha = cFecha.getDateCellValue();
            List<String> personal = new ArrayList<>();
            for (int c = COL_INICIALES; c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String val = cell.getStringCellValue().trim();
                    if (!val.isEmpty()) personal.add(val);
                }
            }
            registros.add(new RegistroFila(fecha, personal));
        }

        registros.sort(Comparator.comparing(r -> r.fecha));
        return registros;
    }

    // Devuelve la lista de trabajadores que había justo antes (o en) la fecha indicada.
    // Sirve para saber con qué plantilla hay que trabajar cuando se inserta un cambio retroactivo.
    private List<String> getPlantillaEnFecha(List<RegistroFila> registros, Date fecha) {
        List<String> resultado = new ArrayList<>();
        for (RegistroFila reg : registros) {
            if (!reg.fecha.after(fecha)) {
                resultado = new ArrayList<>(reg.personal);
            }
        }
        return resultado;
    }

    // ── Reescritura del Excel ────────────────────────────────────────────────

    // Este es el método más importante del programa. En vez de añadir filas al final,
    // borra toda la zona de datos y la vuelve a escribir desde cero con los registros ordenados.
    // Así se evitan duplicados de trimestres y desorden visual.
    private void reescribirExcel(List<RegistroFila> registros) {
        // Primero limpiamos la zona de datos antigua (filas y merges)
        int ultimaFila = hoja.getLastRowNum();
        for (int i = ultimaFila; i >= FILA_DATOS_INI; i--) {
            Row filaAntigua = hoja.getRow(i);
            if (filaAntigua != null) hoja.removeRow(filaAntigua);
        }
        for (int i = hoja.getNumMergedRegions() - 1; i >= 0; i--) {
            if (hoja.getMergedRegion(i).getFirstRow() >= FILA_DATOS_INI) {
                hoja.removeMergedRegion(i);
            }
        }

        int maxPersonal = 0;
        for (RegistroFila r : registros) maxPersonal = Math.max(maxPersonal, r.personal.size());
        int maxCol = Math.max(COL_INICIALES, COL_INICIALES + maxPersonal - 1);

        borrarZonaDatos();
        actualizarFilaNumeracion(maxPersonal);

        // Primera pasada: calculamos en qué filas del Excel va cada año y cada trimestre.
        // Necesitamos saberlo antes de escribir para poder hacer los merges correctamente.
        Map<String, int[]> rangoAÑO       = new LinkedHashMap<>();
        Map<String, int[]> rangoTrimestre  = new LinkedHashMap<>();
        List<String> clavesTrimestre       = new ArrayList<>();
        Map<String, List<Integer>> filasPorTrimestre = new LinkedHashMap<>();

        String AÑOActual = null, trimActual = null;
        int fila = FILA_DATOS_INI;
        for (RegistroFila reg : registros) {
            String AÑO = "AÑO " + getAÑO(reg.fecha);
            String trim = getTrimestre(reg.fecha) + " " + getAÑO(reg.fecha);

            if (!AÑO.equals(AÑOActual)) {
                rangoAÑO.put(AÑO, new int[]{fila, fila});
                AÑOActual = AÑO;
            } else {
                rangoAÑO.get(AÑO)[1] = fila;
            }

            if (!trim.equals(trimActual)) {
                rangoTrimestre.put(trim, new int[]{fila, fila});
                clavesTrimestre.add(trim);
                filasPorTrimestre.put(trim, new ArrayList<>());
                trimActual = trim;
            } else {
                rangoTrimestre.get(trim)[1] = fila;
            }

            filasPorTrimestre.get(trim).add(fila);
            fila++;
        }

        // Segunda pasada: escribimos las filas con sus estilos
        AÑOActual = null;
        trimActual = null;
        fila = FILA_DATOS_INI;

        for (RegistroFila reg : registros) {
            String AÑO = "AÑO " + getAÑO(reg.fecha);
            String trim = getTrimestre(reg.fecha) + " " + getAÑO(reg.fecha);
            boolean nuevoAÑO      = !AÑO.equals(AÑOActual);
            boolean nuevoTrimestre = !trim.equals(trimActual);

            Row row = hoja.createRow(fila);

            // El año solo se escribe en la primera fila del bloque; el resto se une con merge
            if (nuevoAÑO) {
                Cell c = row.createCell(COL_AÑO);
                c.setCellValue(AÑO);
                c.setCellStyle(estiloAÑO);
                AÑOActual = AÑO;
            }

            // El trimestre igual: solo en la primera fila del bloque
            if (nuevoTrimestre) {
                Cell c = row.createCell(COL_TRIMESTRE);
                c.setCellValue(trim);
                boolean unaFila = rangoTrimestre.get(trim)[0] == rangoTrimestre.get(trim)[1];
                c.setCellStyle(unaFila ? estiloTrimestreUnico : estiloTrimestre);
                trimActual = trim;
            }

            Cell cFecha = row.createCell(COL_FECHA);
            cFecha.setCellValue(reg.fecha);
            cFecha.setCellStyle(estiloFecha);

            Cell cNum = row.createCell(COL_NUM_PERS);
            cNum.setCellValue(reg.personal.size());
            cNum.setCellStyle(estiloNumero);

            for (int i = 0; i < reg.personal.size(); i++) {
                Cell c = row.createCell(COL_INICIALES + i);
                c.setCellValue(reg.personal.get(i));
                c.setCellStyle(estiloInicial);
            }

            // Las celdas que no tienen trabajador se pintan de gris para que no queden vacías
            for (int col = COL_INICIALES + reg.personal.size(); col <= maxCol; col++) {
                Cell c = row.createCell(col);
                c.setCellStyle(estiloGris);
            }

            fila++;
        }

        // Merges del bloque AÑO: unimos todas las filas del mismo año en la columna B
        for (Map.Entry<String, int[]> e : rangoAÑO.entrySet()) {
            int[] r = e.getValue();
            if (r[0] == r[1]) {
                aplicarBordesDirectos(r[0], COL_AÑO, true, true, true, false);
            } else {
                CellRangeAddress region = new CellRangeAddress(r[0], r[1], COL_AÑO, COL_AÑO);
                hoja.addMergedRegion(region);
                RegionUtil.setBorderLeft(BorderStyle.THIN, region, hoja);
                RegionUtil.setBorderTop(BorderStyle.THIN, region, hoja);
                RegionUtil.setBorderBottom(BorderStyle.THIN, region, hoja);
            }
        }

        // Merges del bloque TRIMESTRE: unimos las filas en columna C y también en D (% trimestral)
        for (Map.Entry<String, int[]> e : rangoTrimestre.entrySet()) {
            String trimKey = e.getKey();
            int[] r = e.getValue();

            if (r[0] != r[1]) {
                CellRangeAddress regionTrim = new CellRangeAddress(r[0], r[1], COL_TRIMESTRE, COL_TRIMESTRE);
                hoja.addMergedRegion(regionTrim);
                RegionUtil.setBorderRight(BorderStyle.THIN, regionTrim, hoja);
                RegionUtil.setBorderTop(BorderStyle.THIN, regionTrim, hoja);
                RegionUtil.setBorderBottom(BorderStyle.THIN, regionTrim, hoja);

                CellRangeAddress regionPct = new CellRangeAddress(r[0], r[1], COL_PCT_TRIM, COL_PCT_TRIM);
                hoja.addMergedRegion(regionPct);
            }

            // El % trimestral solo se calcula cuando el trimestre ya ha terminado,
            // es decir, cuando existe un trimestre posterior en el Excel.
            // La fórmula usa 20 como media fija (se cambia manualmente a final de año).
            boolean cerrado = !trimKey.equals(clavesTrimestre.get(clavesTrimestre.size() - 1));
            if (cerrado) {
                int maxVal = 0;
                for (int fr : filasPorTrimestre.get(trimKey)) {
                    maxVal = Math.max(maxVal, getNumPersonasFila(fr));
                }
                if (maxVal > 0) {
                    Row firstRow = hoja.getRow(r[0]);
                    if (firstRow != null) {
                        Cell cPct = firstRow.getCell(COL_PCT_TRIM);
                        if (cPct == null) cPct = firstRow.createCell(COL_PCT_TRIM);
                        cPct.setCellFormula("(100-((20*100)/" + maxVal + "))/100");
                        cPct.setCellStyle(estiloPct);
                    }
                }
            }
        }

        // Colores máximo y mínimo: solo para trimestres ya cerrados
        for (int t = 0; t < clavesTrimestre.size() - 1; t++) {
            String trimKey = clavesTrimestre.get(t);
            List<Integer> filas = filasPorTrimestre.get(trimKey);
            if (filas.isEmpty()) continue;

            int maxVal = Integer.MIN_VALUE, minVal = Integer.MAX_VALUE;
            for (int fr : filas) {
                int v = getNumPersonasFila(fr);
                maxVal = Math.max(maxVal, v);
                minVal = Math.min(minVal, v);
            }

            // Verde en la primera fila donde se alcanzó el máximo
            for (int fr : filas) {
                if (getNumPersonasFila(fr) == maxVal) { pintarCeldaF(fr, estiloVerde); break; }
            }

            // Naranja en la última fila donde se alcanzó el mínimo (solo si es distinto al máximo)
            if (minVal != maxVal) {
                int filaMin = -1;
                for (int fr : filas) { if (getNumPersonasFila(fr) == minVal) filaMin = fr; }
                if (filaMin != -1) pintarCeldaF(filaMin, estiloNaranja);
            }
        }
    }

    // Borra filas y merges de la zona de datos (a partir de la fila 5).
    // Las filas del título, numeración y cabeceras no se tocan.
    private void borrarZonaDatos() {
        for (int i = hoja.getNumMergedRegions() - 1; i >= 0; i--) {
            if (hoja.getMergedRegion(i).getFirstRow() >= FILA_DATOS_INI) {
                hoja.removeMergedRegion(i);
            }
        }
        for (int r = hoja.getLastRowNum(); r >= FILA_DATOS_INI; r--) {
            Row row = hoja.getRow(r);
            if (row != null) hoja.removeRow(row);
        }
    }

    // ── Estilos ──────────────────────────────────────────────────────────────

    // Crea todos los estilos que se van a usar al escribir el Excel.
    // Se detecta la fuente del documento original para que los datos nuevos
    // queden visualmente iguales a los que ya había.
    private void inicializarEstilos() {
        Font fDoc = detectarFuenteDocumento();
        short tamano = fDoc.getFontHeightInPoints();
        String nombre = fDoc.getFontName();
        DataFormat fmt = workbook.createDataFormat();

        Font fNormal = workbook.createFont();
        fNormal.setFontName(nombre);
        fNormal.setFontHeightInPoints(tamano);

        Font fNegrita = workbook.createFont();
        fNegrita.setFontName(nombre);
        fNegrita.setFontHeightInPoints(tamano);
        fNegrita.setBold(true);

        estiloAÑO = workbook.createCellStyle();
        estiloAÑO.setFont(fNegrita);
        estiloAÑO.setAlignment(HorizontalAlignment.CENTER);
        estiloAÑO.setVerticalAlignment(VerticalAlignment.CENTER);
        aplicarBordesFinos(estiloAÑO);

        estiloTrimestre = workbook.createCellStyle();
        estiloTrimestre.setFont(fNormal);
        estiloTrimestre.setAlignment(HorizontalAlignment.CENTER);
        estiloTrimestre.setVerticalAlignment(VerticalAlignment.CENTER);
        aplicarBordesFinos(estiloTrimestre);

        // Cuando un trimestre solo tiene una fila, no se puede usar merge,
        // así que los bordes se aplican directamente en el estilo
        estiloTrimestreUnico = workbook.createCellStyle();
        estiloTrimestreUnico.setFont(fNormal);
        estiloTrimestreUnico.setAlignment(HorizontalAlignment.CENTER);
        estiloTrimestreUnico.setVerticalAlignment(VerticalAlignment.CENTER);
        aplicarBordesFinos(estiloTrimestreUnico);

        estiloPct = workbook.createCellStyle();
        estiloPct.setFont(fNormal);
        estiloPct.setDataFormat(fmt.getFormat("0.00%"));
        estiloPct.setAlignment(HorizontalAlignment.CENTER);
        estiloPct.setVerticalAlignment(VerticalAlignment.CENTER);
        aplicarBordesFinos(estiloPct);

        estiloFecha = workbook.createCellStyle();
        estiloFecha.setFont(fNormal);
        estiloFecha.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
        estiloFecha.setAlignment(HorizontalAlignment.CENTER);
        estiloFecha.setVerticalAlignment(VerticalAlignment.CENTER);
        aplicarBordesFinos(estiloFecha);

        estiloNumero = workbook.createCellStyle();
        estiloNumero.setFont(fNormal);
        estiloNumero.setAlignment(HorizontalAlignment.CENTER);
        estiloNumero.setVerticalAlignment(VerticalAlignment.CENTER);
        aplicarBordesFinos(estiloNumero);

        estiloInicial = workbook.createCellStyle();
        estiloInicial.setFont(fNormal);
        estiloInicial.setAlignment(HorizontalAlignment.CENTER);
        estiloInicial.setVerticalAlignment(VerticalAlignment.CENTER);
        aplicarBordesFinos(estiloInicial);

        estiloGris = workbook.createCellStyle();
        ((XSSFCellStyle) estiloGris).setFillForegroundColor(toXSSFColor(COLOR_GRIS));
        estiloGris.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        aplicarBordesFinos(estiloGris);

        estiloVerde = workbook.createCellStyle();
        estiloVerde.setFont(fNormal);
        estiloVerde.setAlignment(HorizontalAlignment.CENTER);
        ((XSSFCellStyle) estiloVerde).setFillForegroundColor(toXSSFColor(COLOR_VERDE));
        estiloVerde.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        aplicarBordesFinos(estiloVerde);

        estiloNaranja = workbook.createCellStyle();
        estiloNaranja.setFont(fNormal);
        estiloNaranja.setAlignment(HorizontalAlignment.CENTER);
        ((XSSFCellStyle) estiloNaranja).setFillForegroundColor(toXSSFColor(COLOR_NARANJA));
        estiloNaranja.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        aplicarBordesFinos(estiloNaranja);
    }

    // Pone borde fino en los cuatro lados de una celda
    private void aplicarBordesFinos(CellStyle st) {
        st.setBorderTop(BorderStyle.THIN);
        st.setBorderBottom(BorderStyle.THIN);
        st.setBorderLeft(BorderStyle.THIN);
        st.setBorderRight(BorderStyle.THIN);
    }

    // Lee las celdas de datos existentes para detectar qué fuente usa el documento.
    // Si no hay datos previos, usa Calibri 10 por defecto.
    private Font detectarFuenteDocumento() {
        for (int r = FILA_DATOS_INI; r <= hoja.getLastRowNum(); r++) {
            Row row = hoja.getRow(r);
            if (row == null) continue;
            for (int col = COL_FECHA; col <= COL_INICIALES + 3; col++) {
                Cell c = row.getCell(col);
                if (c != null && c.getCellType() != CellType.BLANK) {
                    Font f = workbook.getFontAt(c.getCellStyle().getFontIndex());
                    if (f != null && f.getFontHeightInPoints() > 0) return f;
                }
            }
        }
        Font fallback = workbook.createFont();
        fallback.setFontName("Calibri");
        fallback.setFontHeightInPoints((short) 10);
        return fallback;
    }

    // ── Numeración de posiciones ─────────────────────────────────────────────

    // Actualiza la fila de numeración (fila 3) si hay más trabajadores que nunca antes.
    // Por ejemplo si hay 25 trabajadores y solo había números hasta el 24, añade el 25.
    private void actualizarFilaNumeracion(int maxPersonal) {
        Row filaNums = hoja.getRow(FILA_NUMEROS);
        if (filaNums == null) filaNums = hoja.createRow(FILA_NUMEROS);
        int maxActual = 0;
        for (Cell c : filaNums) {
            if (c.getCellType() == CellType.NUMERIC)
                maxActual = Math.max(maxActual, (int) c.getNumericCellValue());
        }
        if (maxPersonal <= maxActual) return;
        CellStyle stNum = workbook.createCellStyle();
        stNum.setAlignment(HorizontalAlignment.CENTER);
        for (int i = maxActual + 1; i <= maxPersonal; i++) {
            Cell c = filaNums.createCell(COL_INICIALES + i - 1);
            c.setCellValue(i);
            c.setCellStyle(stNum);
        }
    }
    

    // ── Métodos auxiliares ───────────────────────────────────────────────────

    private void pintarCeldaF(int fila, CellStyle estilo) {
        Row row = hoja.getRow(fila);
        if (row == null) return;
        Cell c = row.getCell(COL_NUM_PERS);
        if (c != null) c.setCellStyle(estilo);
    }

    private void aplicarBordesDirectos(int fila, int col,
                                        boolean arriba, boolean abajo,
                                        boolean izquierda, boolean derecha) {
        Row row = hoja.getRow(fila);
        if (row == null) return;
        Cell c = row.getCell(col);
        if (c == null) return;
        XSSFCellStyle st = (XSSFCellStyle) workbook.createCellStyle();
        st.cloneStyleFrom(c.getCellStyle());
        if (arriba)    st.setBorderTop(BorderStyle.THIN);
        if (abajo)     st.setBorderBottom(BorderStyle.THIN);
        if (izquierda) st.setBorderLeft(BorderStyle.THIN);
        if (derecha)   st.setBorderRight(BorderStyle.THIN);
        c.setCellStyle(st);
    }

    private int getNumPersonasFila(int fila) {
        Row row = hoja.getRow(fila);
        if (row == null) return 0;
        Cell c = row.getCell(COL_NUM_PERS);
        if (c == null || c.getCellType() != CellType.NUMERIC) return 0;
        return (int) c.getNumericCellValue();
    }

    private XSSFColor toXSSFColor(byte[] rgb) {
        return new XSSFColor(new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF), null);
    }
    

    String getTrimestre(Date fecha) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(fecha);
        int mes = cal.get(Calendar.MONTH) + 1;
        if (mes <= 3) return "1T";
        if (mes <= 6) return "2T";
        if (mes <= 9) return "3T";
        return "4T";
    }

    int getAÑO(Date fecha) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(fecha);
        return cal.get(Calendar.YEAR);
    }

    Date parsearFecha(String fechaStr) throws ParseException {
        if (fechaStr == null || fechaStr.isEmpty()) return new Date();
        FMT.setLenient(false);
        try {
            return FMT.parse(fechaStr);
        } catch (ParseException e) {
            throw new ParseException("Formato de fecha inválido. Usa dd/MM/yyyy (ej: 15/04/2026)", 0);
        }
    }

    private void guardar(String ruta) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(ruta)) {
            workbook.write(fos);
        }
        System.out.println("  Guardado en : " + ruta);
    }

    // Convierte un color en formato hexadecimal (ej: "d8d8d8") a un array de bytes RGB
    private static byte[] hexToRgb(String hex) {
        int r = Integer.parseInt(hex.substring(0, 2), 16);
        int g = Integer.parseInt(hex.substring(2, 4), 16);
        int b = Integer.parseInt(hex.substring(4, 6), 16);
        return new byte[]{(byte) r, (byte) g, (byte) b};
    }
    public String getPlantillaComoTexto(String rutaExcel) {
    try (FileInputStream fis = new FileInputStream(rutaExcel);
         Workbook tempWorkbook = new XSSFWorkbook(fis)) {
        
        Sheet hoja = tempWorkbook.getSheetAt(0);
        int ultimaFila = hoja.getLastRowNum();
        
        // Buscamos la última fila que tenga datos de personas (Columna F / índice 5)
        Row filaDatos = null;
        for (int i = ultimaFila; i >= 0; i--) {
            Row r = hoja.getRow(i);
            if (r != null && r.getCell(5) != null && r.getCell(5).getCellType() == CellType.NUMERIC) {
                filaDatos = r;
                break;
            }
        }

        if (filaDatos == null) return "No se encontraron datos en el Excel.";

        StringBuilder sb = new StringBuilder();
        sb.append("ESTADO ACTUAL DE LA PLANTILLA:\n");
        sb.append("--------------------------------------------------\n");

        int contador = 1;
        // Las iniciales empiezan en la columna G (índice 6) en adelante
        for (int c = 6; c < filaDatos.getLastCellNum(); c++) {
            Cell celda = filaDatos.getCell(c);
            if (celda != null && celda.getCellType() == CellType.STRING && !celda.getStringCellValue().trim().isEmpty()) {
                sb.append(String.format("[%d] %s\n", contador++, celda.getStringCellValue()));
            }
        }
        
        sb.append("--------------------------------------------------\n");
        sb.append("Total: " + (contador - 1) + " personas.");
        return sb.toString();

    } catch (Exception e) {
        return "Error al leer el archivo: " + e.getMessage();
    }
}
}

