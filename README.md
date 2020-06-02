# AddPageNumber to PDF

**Aplicación HTA que permite añadir números de página a un fichero PDF**

## Consideraciones iniciales

Existen dos implementaciones diferentes de la herramienta, una hace uso de una intefaz HTA (IE9) y otra hace uso del objeto WScript y se ejecuta por línea de comandos o mediante paso de argumentos arrastrando en el explorador de ficheros, el archivo PDF sobre el script.


## Requisitos previos

* Requiere tener instalada la aplicación PDFCreator v.1.7.2 (no testeado en versiones posteriores)
* Requiere el plugin pdfforge (incluido en la versión de PDFCreator 1.7.2)

## Uso

***Versión HTA***

Ejecutar el fichero **AddPageNumber.hta**

Abrirá la ventana de la aplicación con los parámtros configurables para añadir la numeración de páginas. El archivo paginado resultante se guarda en la misma ubicación del archivo original con el añadido ´_pag´ al final del nombre

***Versión WSCript***

Versión línea de comandos o explorador de archivos. Mediante línea de comandos se ejecuta el script pasándole como parámetro el fichero PDF a paginas:

```c:\ wscript.exe AddPageNumber.vbs <FICHERO.PDF>```

Para hacer lo mismo a través del explorador de archivos de Windows, arrastra el fichero PDF a páginar y suéltalo sobre el fichero AddPageNumber.vbs.

El archivo paginado resultante se guarda en la misma ubicación del archivo original con el añadido ´_pag´ al final del nombre


## Versiones

Versión WScript v1.0.0 12/03/2020
Versión HTA 	v.1.10 01/06/2020
## Autor

* **Miguel Angel Camacho** - [miguelangelcamacho.com]

## Licencia

Ver [LICENCIA]

[miguelangelcamacho.com]:https://www.miguelangelcamacho.com
[LICENCIA]:LICENSE