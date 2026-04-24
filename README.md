# Gestión de Plantilla - Rotación de Personal

Aplicación de escritorio desarrollada en Java (JavaFX) para la gestión de la rotación de personal en una plantilla de trabajo. El sistema permite registrar entradas y salidas de trabajadores y actualizar automáticamente un archivo Excel con la información organizada por fechas y periodos.

El objetivo del programa es sustituir la edición manual del Excel por un sistema automatizado, rápido y menos propenso a errores.

---

## Tecnología utilizada

- Java 21 (LTS)
- JavaFX
- Maven
- jpackage (para generar ejecutable nativo de Windows)

---

## Ejecutable (IMPORTANTE)

Este proyecto ya no se ejecuta desde consola.

Se entrega como aplicación de escritorio:

- Ejecutable generado con jpackage
- Tipo: app-image (aplicación portable)
- No requiere instalación de Java en el equipo destino

### Ejecución

Abrir el archivo:

GestionPlantilla.exe

---

## Estructura del proyecto

GestionPlantillaFinal/
│
├── src/                  Código fuente Java
├── pom.xml              Configuración Maven
├── target/              Archivos generados (no incluir en GitHub)
├── dist_prod/           Ejecutable generado (jpackage)

---

## Compilación del proyecto (solo desarrollo)

Para compilar el proyecto:

mvn clean package

Genera el archivo .jar en la carpeta target/.

---

## Generación del ejecutable

El ejecutable se genera con jpackage usando tipo:

--type app-image

Esto crea una aplicación independiente que incluye:

- Ejecutable .exe
- Runtime de Java incluido
- Librerías necesarias
- Estructura completa de ejecución

---

## Funcionamiento de la aplicación

Al iniciar la aplicación:

1. Se abre la interfaz gráfica (JavaFX)
2. El usuario selecciona el archivo Excel
3. El sistema guarda la ruta automáticamente

---

## Funcionalidades principales

- Registro de entrada de trabajadores
- Registro de salida de trabajadores
- Visualización de plantilla actual
- Cambio de archivo Excel
- Actualización automática del Excel

---

## Funcionamiento interno del Excel

Cada cambio en la plantilla:

1. Lee el Excel completo
2. Añade el nuevo registro
3. Ordena los datos cronológicamente
4. Reescribe el archivo desde cero

Esto garantiza:

- Orden automático por fechas
- Agrupación por año y trimestre
- Cálculo automático de porcentajes trimestrales
- Formato visual estructurado

---

## Notas técnicas

- La aplicación es independiente (no requiere Java instalado)
- El runtime de Java está incluido en el ejecutable
- Compatible con Windows
- El archivo de configuración se guarda en:

Documents/GestionPlantilla/config.properties

---

## Autor

Proyecto desarrollado como aplicación de gestión interna para automatización de procesos administrativos.
