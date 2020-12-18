
PROJECT

Su objetivo es remitir a múltiples direcciones una plantilla "Email.htm", que puede ser editada por 
ejemplo, fácilmente, a través de Microsoft Word. En ella hay corchetes ([]) que se sustituirán por los 
elementos del archivo "Destinatarios.xlsx", del tercero en adelante (la primera y segunda columna 
corresponden a los emails de los destinatarios en copia separados por comas y al asunto de cada email).

Se deben crear tantas carpetas en la carpeta "Adjuntos", como emails se deseen remitir, una por cada
fila del archivo "Destinatarios.xlsx", que se identificarán en orden descendente por nombre con dichas
filas. Si denominásemos las carpetas como "Archivo" y un número, "Archivo1" corresponderá a la primera 
fila, "Archivo2" a la segunda... Cada carpeta dentro de "Adjuntos" contendrá los archivos adjuntos que
serán enviados a cada conjunto de destinatarios de cada fila del archivo "Destinatarios.xlsx".

Los datos de fechas y números en la hoja EXCEL de Destinatarios no se reflejan adecuadamente en el
mensaje email.htm.

El programa está pensado para Windows y se ejecuta desde "PMEmails.bat", aunque puede modificarse
sencillamente para otros sistemas operativos.

Para que funcione el programa debe estar instalado Python 3.7 al menos, además del paquete beautifulsoup4 
y python-docx (el archivo "PMEmails.bat" los instala antes de ejecutar el programa principal).

MODULES

El fichero batch ejecuta PMEmails.py que es el único módulo del programa y el archivo Config.json tan
sólo sirve para guardar el servidor de correo y dirección email del último uso del programa.

CONTACT AND CONTRIBUTION

Este es un proyecto personal que no está abierto a contribuciones ahora mismo pero siéntete libre de compartir
cualquier comentario, pregunta o sugerencia a través de javiercuartasmicieces@hotmail.com.

ACKNOWLEDGEMENTS

Los scripts fueron desarrollados utilizando las librerías beautifulsoup4 y python-docx.
