@echo off
call %userprofile%\Anaconda3\Scripts\activate.bat
pip install beautifulsoup4
pip install python-docx
SET ruta0=%cd%\PMEmails.py
%userprofile%\Anaconda3\python.exe %ruta0%