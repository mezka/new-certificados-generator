# Script interno para generación masiva de certificados de puertas
A partir de templates HTML, y archivos PDF prerenderizados, con impresión de marca de agua

## Instalación de dependencias en Windows

Para utilizar la aplicación se deben tener instalados los siguientes paquetes:

* Python 3 (Probado con Python 3.8.3 x64)
* ImageMagick (Probado con v7.0.10-22 x64 dinamica con 16 bits por pixel)
* Runtime GTK (Probado con v3.24.18 x64)

Los paquetes mencionados anteriormente pueden ser obtenidos de:

https://www.python.org/downloads/

https://www.imagemagick.org/script/download.php

https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer

O de los siguientes mirror proporcionados por el autor:

[Python 3.8.3](https://1drv.ms/u/s!Arz535PAeGSPjFLSB01egBbotpyA?e=RM87ef)

[ImageMagick v7.0.10-22 x64 dinamica con 16 bits por pixel](https://1drv.ms/u/s!Arz535PAeGSPjFMIAK43ABJK08Ky?e=ZhBerz)

[Runtime GTK v3.24.18 x64](https://1drv.ms/u/s!Arz535PAeGSPjFSDIBOU2qDkxuxl?e=OCTubg)

## Instalacion de dependencias en distribuciones de Linux basadas en Debian

```bash
sudo apt-get update
sudo apt-get install libffi-dev libxml2-dev libxslt1-dev
```

Además se debe configurar la libreria externa ImageMagick para habilitar su uso para leer y escribir archivos en formato PDF modificando la siguiente linea en `/etc/ImageMagick-*/policy.xml`

De:
```bash
<policy domain="coder" rights="none" pattern="PDF" />
```

A:
```bash
<policy domain="coder" rights="read|write" pattern="PDF" />
```

De no realizar este paso es probable que el script falle con el siguiente error:

```bash
wand.exceptions.PolicyError: not authorized `temp/temp_watermark.pdf' @ error/constitute.c/WriteImage/1037
```

## Instalación de los modulos de Python usados por el script

Para instalar los modulos usados por el script:

```bash
pip install -r requirements.txt
```

## Utilización del script

Para cargar los datos del script se utiliza la planilla de calculo `datos.xlsx`, la misma tiene la siguiente estructura por defecto:

| CLIENTE  | DIRECCION_ENTREGA| MODELO  | TIPO_DE_FACTURA | FACTURA_NRO | REMITO_NRO | MEDIDA   | CANTIDAD | NROS_SERIE |
|:--------:|:----------------:|:-------:|:---------------:|:-----------:|:----------:|:--------:|:--------:|:----------:|
| MESQUITA | Alsina 2501      | RF60    | A               | 236         | 289        | 900x2000 | 2        | 200, 201   |
| PELIKAN  | Sarmiento 1232   | RF30    | C               | 237         | 290        | 800x2000 | 1        | 202        |

Se generara un certificado por cada fila ingresada en la planilla de calculo. A partir del codigo de modelo ingresado en la columna MODELO el script elegira el certificado a incluir, los modelos disponibles son los siguientes:

| CODIGO DE MODELO           |        CERTIFICADO A UTILIZAR                             |
|:--------------------------:|:---------------------------------------------------------:|
| RF30	                     |  Incluye el certificado RF30 hoja simple                  |	   
| RF60_SIMPLE	             |  Incluye el certificado RF60 hoja simple                  |	   
| RF60_DOBLE	             |  Incluye el certificado RF60 hoja doble	                 |
| RF90	                     |  Incluye el certificado RF120 hoja simple	             |
| RF120	                     |  Incluye el certificado RF120 hoja simple	             |
| RF60_EURO_SIMPLE	         |  Incluye el certificado RF60 hoja simple europea	         |  
| RF60_EURO_DOBLE	         |  Incluye el certificado RF60 hoja doble europea	         |
| RF60_EURO_SIMPLE_VIDRIADA	 |  Incluye el certificado RF60 hoja simple europea vidriada |	   
| RF120_EURO_SIMPLE	         |  Incluye el certificado RF120 hoja simple europea	     |
| RF120_EURO_DOBLE  	     |  Incluye el certificado RF120 hoja doble europea	         |
| PORTON_RF60_CORREDIZO      |  Incluye el certificado RF60 para porton corredizo        |	   

Una vez cargados los datos de los certificados a generar abrimos una terminal en el directorio del script, el mismo tiene la siguiente estructura:

```
out
pdf_files
templates
config.py
datos.xlsx
main.py
README.md
requirements.txt
```

Una vez situados en este directorio ejecutaremos el siguiente comando:

```bash
python main.py
```

Esperamos un momento y encontraremos los certificados generados en la carpeta `out`