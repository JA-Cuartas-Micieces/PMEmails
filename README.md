
PROJECT

The goal of this project is sending emails to several addresses according to a template "Email.htm".
This template can be edited using Microsoft Word or any other text editor. The brackets positions 
in the template will be replaced by "Destinatarios.xlsx" elements. The first column contains
email recipients separated by semicolon and the second column contains the email subject for each 
email.

You must create as many folders as emails you want to send inside the "Adjuntos" folders. There must 
be at least one folder for each row of "Destinatarios.xlsx". Folders are sorted alphabetically in 
descending order, for rows and attachments folders' to match.

The batch file ("PMEmails.bat") is made to run the script on Windows, though it can be replaced by a
basch script or other depending on the OS.

The script is made for Python 3.7 (Anaconda) and beautifulsoup4 and python-docx libraries must be 
installed for the script to work. "PMEmails.bat" installs both packages if needed.


MODULES

The batch file runs PMEmails.py the only script module. Config.json saves the email server and the
email address from the last time the script is run.

CONTACT AND CONTRIBUTION

This is a personal project and It is not open to contributions right now though I am open to it, 
feel free to share any comment, question or suggestion through javiercuartasmicieces@hotmail.com.

ACKNOWLEDGEMENTS

The script was developed using beautifulsoup4 and python-docx libraries.
