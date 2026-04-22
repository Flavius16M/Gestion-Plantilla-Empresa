# Gestión de Plantilla - Rotación de Personal

Programa de consola en Java para registrar entradas y salidas de trabajadores en el Excel de rotación de personal del departamento de Calidad. El objetivo es que esto se haga en segundos en lugar de modificar el Excel a mano.

---

## Requisitos

- Java 11 o superior
- Maven 3.x (para compilar)

Para comprobar que están instalados:
```
java -version
mvn -version
```

---

## Compilar el proyecto

Desde la carpeta raíz del proyecto (donde está el `pom.xml`):

```
mvn package
```

Esto genera el archivo `target/GestionPlantilla.jar` con todas las dependencias incluidas.

---

## Ejecutar el programa

En Windows, hacer doble clic en `ejecutar.bat`. Este archivo configura la consola para que los caracteres especiales (tildes, símbolos del menú) se vean bien.

O desde la terminal:
```
java -jar target/GestionPlantilla.jar
```

---

## Cómo funciona

Al abrirlo la primera vez, el programa pide seleccionar el archivo Excel mediante una ventana del explorador de Windows (así no hay que escribir la ruta a mano y no hay problemas con nombres que tienen tildes). Esa ruta se guarda y no vuelve a pedirse.

A partir de ahí, el menú tiene estas opciones:

**1 - Entrada de trabajador**
Pide las iniciales y la fecha del cambio. Si no se pone fecha, usa la de hoy. Añade una nueva fila al Excel con el trabajador incorporado al final de la lista.

**2 - Salida de trabajador**
Muestra la plantilla actual, pide las iniciales del que se va y la fecha. Elimina al trabajador y compacta la lista para que no queden huecos.

**3 - Ver plantilla actual**
Muestra por consola quién está en la plantilla en este momento sin modificar nada.

**4 - Cambiar archivo Excel**
Por si se mueve el archivo o se pasa el programa a otro ordenador.

---

## Cómo actualiza el Excel

Cada vez que se registra un cambio, el programa:

1. Lee todos los registros existentes del Excel
2. Añade el nuevo registro
3. Ordena todo por fecha
4. Borra la zona de datos y la reescribe desde cero

Esto hace que el Excel siempre quede ordenado cronológicamente aunque se introduzcan fechas antiguas o fuera de orden. Los bloques de año y trimestre se fusionan solos, se pintan las celdas de máximo (verde) y mínimo (naranja) de cada trimestre cerrado, y se calcula el % trimestral automáticamente.

---

## Estructura del Excel generado

| Col B    | Col C      | Col D       | Col E  | Col F      | Col G, H... |
|----------|------------|-------------|--------|------------|-------------|
| AÑO 2026 | 2T 2026    | % trimest.  | Fecha  | Nº personas | MA, GC...  |
|          |            |             |        |            |             |

- El año y el trimestre se fusionan verticalmente en todas sus filas
- El % trimestral aparece cuando el trimestre ya ha terminado. El 20 de la fórmula es la media fija y se cambia a mano al cerrar el año
- Las celdas sin trabajador se rellenan de gris (#d8d8d8)

---

## Notas

- El programa no borra el historial, solo añade registros nuevos y reordena
- Las iniciales no distinguen mayúsculas/minúsculas al buscar
- El archivo de configuración se guarda en `Documents/GestionPlantilla/config.properties`
