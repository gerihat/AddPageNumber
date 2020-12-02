# AddPageNumber to PDF

**HTA Application to add page numaer to a given PDF file**

## Initial considerations

There are two different implementations, one use HTA interface (IE9) and the second use WSscript object in command line execution or drag & drop PDF file over script file.

## Prerequisites

* PDFCreator v.1.7.2 (non-tested later versions)
* Pdfforge plugin (included with PDFCreator 1.7.2 version)

## How to use

***HTA Version***

Run **AddPageNumber.hta**

App window GUI will opens with options panel to configure the numeration and the numerated PDF file will be saved in the same location with ´_pag´ added to the name.

***WSCript Version***

Command line version or File Explorer drag & drop. 

Command line script:
```c:\ wscript.exe AddPageNumber.vbs <Filename.PDF>```

Also in the Windows File Explorer you can drag & drop the PDF file on AddPageNumer.vbs script

The numerated PDF file will be saved in the same location with ´_pag´ added to the name.


## Versions

Version WScript v1.0.0 12/03/2020
Version HTA 	v.1.10 01/06/2020

## Autor

* **Miguel Angel Camacho** - [miguelangelcamacho.com]

## License

See [LICENSE]

[miguelangelcamacho.com]:https://www.miguelangelcamacho.com
[LICENSE]:LICENSE
