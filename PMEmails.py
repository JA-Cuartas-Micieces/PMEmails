# -*- coding: utf-8 -*-
"""

Created on Wed Aug 19 20:47:00 2020

Copyright 2020 Javier A. Cuartas Micieces

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in 
the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of 
the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS 
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER 
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

"""

import xlrd
import json
import subprocess as sp
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import os
import shutil
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import smtplib
from bs4 import BeautifulSoup
import win32com.client as win32
import win32com.client
import os

class TkWindow:
    def __init__(self,*args,**kwargs):
        self.r=Execution()
        self.main_menu()
        
    def main_menu(self):
        self.root = Tk()
        self.root.resizable(0,0)
        self.root.title('Enviar emails')
        self.style=ttk.Style()
        self.fontsizeB='11'
        self.style.configure("TButton",font = ('Arial' , self.fontsizeB))
        
        self.frame_header=ttk.Frame()
        self.frame_header.pack()
        self.frame_body=ttk.Frame()
        self.frame_body.pack()
        self.frame_bottom=ttk.Frame()
        self.frame_bottom.pack()
        
        msg=Message(self.frame_header,text="Por favor, meta en la carpeta Adjuntos, los que desee remitir, rellene el archivo .xlsx en la primera columna con emails de destinatarios separados por comas, segunda con el asunto y el resto con elementos a sustituir por orden, coincidentes con el número de parejas de corchetes ([]) en el archivo Email.htm que será remitido. Después rellene los campos del formulario y pulse enviar para remitir el email dinámico a los destinatarios del archivo .xlsx.",justify=LEFT)
        msg.grid(row=0,column=0,padx=1,pady=1)
        
        entry1=StringVar()
        self.lab1=ttk.Label(self.frame_body,text="Email: ",font = ('Arial' , self.fontsizeB))
        self.lab1.grid(row=1,column=0,padx=(5,0),pady=(5,0))
        self.entry1=ttk.Entry(self.frame_body,width=24,textvariable=entry1)
        self.entry1.insert(0, self.r.fromaddr.encode('UTF-8'))
        self.entry1.grid(row=1,column=1,pady=(5,0))
        entry2=StringVar()
        self.lab2=ttk.Label(self.frame_body,text="Contraseña: ",font = ('Arial' , self.fontsizeB))
        self.lab2.grid(row=2,column=0,padx=(5,0),pady=(5,0))
        self.entry2=ttk.Entry(self.frame_body,width=24,textvariable=entry2, show="*")
        self.entry2.grid(row=2,column=1,pady=(5,0))
        entry3=StringVar()
        self.lab3=ttk.Label(self.frame_body,text="Servidor: ",font = ('Arial' , self.fontsizeB))
        self.lab3.grid(row=3,column=0,padx=(5,0),pady=(5,0))
        self.entry3=ttk.Entry(self.frame_body,width=24,textvariable=entry3)
        self.entry3.insert(0, self.r.nserver.encode('UTF-8'))
        self.entry3.grid(row=3,column=1,pady=(5,0))
        
        bt1path0=self.r.cd+"Email/Email.htm"
        bt1path1=self.r.cd+"Adjuntos"
        bt2path0= self.r.cd+"Destinatarios.xlsx"

        self.bt1=ttk.Button(self.frame_bottom,text="Revisar Email", command=lambda: (self.resetwordfile(bt1path0),self.openword(bt1path0),sp.Popen(f'Explorer {os.path.realpath(bt1path1)}')),style="TButton")
        self.bt1.grid(row=4,column=0,padx=(5,0),pady=(5,5))
        self.bt2=ttk.Button(self.frame_bottom,text="Revisar Destinatarios", command=lambda: self.openexcel(bt2path0),style="TButton")
        self.bt2.grid(row=4,column=1,padx=(5,0),pady=(5,5))
        self.bt3=ttk.Button(self.frame_bottom,text="Enviar", command=lambda: (self.r.send_emails(entry1.get(),entry2.get(),entry3.get())),style="TButton")
        self.bt3.grid(row=4,column=2,padx=(5,5),pady=(5,5))

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()
        
    def on_closing(self):
        self.root.destroy()
        
    def openword(self,path):
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = True
        doc = word.Documents.Open(path)
    
    def resetwordfile(self,path):
        os.chdir(self.r.cd+"\\Email")
        flst=[el for el in os.listdir() if el.endswith("Email.htm")]
        os.chdir(self.r.cd)
        if len(flst)==0:
            document = Document()
            document.save(path)
    
    def openexcel(self,path):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(path)
        
        
class Execution:
        
    @staticmethod
    def appendlist(inp,out):
        for el in inp:
            out.append(el)
            
    @staticmethod
    def extract(a1,a2,strin):
        try:
            sr0=strin
            ia1=sr0.index(a1)+len(a1)
            ia2=sr0[ia1:].index(a2)+ia1
            return sr0[ia1:ia2], ia2+len(a2)
        except:
            return 'EOF',-1
        
    @staticmethod
    def fullextract(a1,a2,html):
        output=''
        sr1=''
        sr0=html
        l=[]
        while(sr0!=sr1):
            sr1=sr0
            output,iax=Execution.extract(a1,a2,sr0)
            if output=='EOF':
    	        break
            l.append(output)
            sr0=sr0[iax:]
        return l
    
    @staticmethod
    def extractall(html,route):
        results=dict()
        for element in route:
            results[element]=Execution.deepextract(html,route,element)
        return results
    
    @staticmethod
    def deepextract(html,route,delimkey):
        delimkeyresults=list()
        if len(route[delimkey])>1:
            D=dict()
            D[route[delimkey][0][0]]=dict()
            D[route[delimkey][0][0]][0]=Execution.fullextract(route[delimkey][0][0],route[delimkey][0][1],html)
        else:
            delimkeyresults.extend(Execution.fullextract(route[delimkey][0][0],route[delimkey][0][1],html))
        for x in range(1,len(route[delimkey])):
            D[route[delimkey][0][0]][x]=list()
            for g in D[route[delimkey][0][0]][x-1]:
                s=Execution.fullextract(route[delimkey][x][0],route[delimkey][x][1],g)
                Execution.appendlist(s,D[route[delimkey][0][0]][x])
            if x==len(route[delimkey])-1:
                delimkeyresults.extend(D[route[delimkey][0][0]][x])
                del D
        return delimkeyresults
    
    @staticmethod
    def replace(strbase,strrep,a,b,e):
        try:
            ia=strbase.find(a)
            sa=strbase[ia+len(a):]
            ib=sa.find(b)
            sb=sa[ib+1:]
            iab=ia+len(a)-1
            s=strbase[:iab]+strrep+sb
            sk=strbase[:ia+len(a)-1]+sa[:ib+len(b)-1]
            return s,sb,iab,sk
        except:
            return 'EOF','EOF','EOF','EOF'
        
    @staticmethod
    def marker_replace(strbase,strrep,a,b):
        try:
            ia=strbase.find(a)
            iab=ia+len(a)-1
            sa=strbase[ia+len(a):]
            ib=sa.find(b)
            sb=sa[ib+len(b):]
            s=strbase[:ia]+strrep+sb
            sk=strbase[:ia]+sa[:ib+len(b)]
            return s,sb,iab,sk
        except:
            return 'EOF',-1
        
    @staticmethod
    def nsubstitute(strbase,lstrrep,nsubstitutions,isubstitutions,a,b,e):
        s=strbase
        sb=strbase
        sr=""
        for i in range(nsubstitutions):
            s,sb,iab,sk=Execution.replace(sb,lstrrep[i],a,b,e)
            if i in isubstitutions:
                sr=sr+s[:iab]+lstrrep[i]
            else:
                sr=sr+sk
        sr=sr+sb
        return sr
    
    @staticmethod
    def marker_nsubstitute(strbase,lstrrep,nsubstitutions,isubstitutions,a,b):
        s=strbase
        sb=strbase
        sr=""
        for i in range(nsubstitutions):
#            print(i)
#            print(lstrrep[i])
            s,sb,iab,sk=Execution.marker_replace(sb,lstrrep[i],a,b)
            if i in isubstitutions:
                sr=sr+s[:iab]+lstrrep[i]
            else:
                sr=sr+sk
        sr=sr+sb
        return sr

    def __init__(self):
        self.cd = str.replace(self.f.__code__.co_filename,"\\","\\\\")[:-11]
        os.chdir(self.cd)
        self.emptyflag="-999999999999999bwiu3"   
        insertv=self.get_insertv()
        self.fromaddr=insertv["Email"]
        self.passw=insertv["Password"]
        self.nserver=insertv["Server"]
        
    def f(self): pass

    def get_insertv(self):
        with open(self.cd+"Config.json") as f:
            inputdict=json.load(f)        
        return inputdict
    
    def send_emails(self,*args,**kwargs):
        try:
            self.fromaddr=args[0]
            self.passw=args[1]
            self.nserver=args[2]
            filename=self.cd+"\\Email\\Email.htm"

            inputWorkbook=xlrd.open_workbook(self.cd+"\\Destinatarios.xlsx")
            bk=inputWorkbook.sheet_by_index(0)
            for v in range(1,bk.nrows):
                with open(filename,"rb") as ht:
                    html=ht.read()

                soup = BeautifulSoup(html,features="lxml")
                encoding = soup.original_encoding or 'utf-8'

                with open(self.cd+"\\Emailm.txt","wb+") as fl:
                    fl.write(soup.encode(encoding))
                with open(self.cd+"\\Emailm.txt","r",encoding=encoding) as fr:
                    self.s0=fr.read()

                os.remove(self.cd+"\\Emailm.txt")
                
                lemailsubstitutions=[]
                for w in range(bk.ncols):
                    val=bk.cell(v, w).value
                    lemailsubstitutions.append(val)

                if len(lemailsubstitutions)-2==len([el for el in Execution.extractall(self.s0,{"[]":[["[","]"]]})["[]"] if all([el=="","if" not in el])]):
                    msg = MIMEMultipart()
                    msg['From'] = self.fromaddr
                    msg['To'] = lemailsubstitutions[0]
                    msg['Subject'] = lemailsubstitutions[1]
                    
                    
                    adjsf=sorted([el for el in os.listdir(self.cd+"\\Adjuntos")])
                    adjsf.reverse()
                    adjs=[el for el in os.listdir(self.cd+"Adjuntos\\"+adjsf[v-1])]
                    for adjn in adjs:
                        adjpath = self.cd+"Adjuntos\\"+adjsf[v-1]+"\\"+adjn
                        adj = open(adjpath, 'rb')
                        adj_MIME = MIMEBase('application', 'octet-stream')
                        adj_MIME.set_payload((adj).read())
                        encoders.encode_base64(adj_MIME)
                        adj_MIME.add_header('Content-Disposition', "attachment; filename= %s" % adjn)
                        msg.attach(adj_MIME)

                    strbase=self.s0
                    a="src=\""
                    b="\""
                    imgtags=Execution.extractall(strbase,{"img":[["<img","/>"]]})["img"]
                    imgs=Execution.extractall(strbase,{"src":[[a,b]]})["src"]
                    lsubs=[]
                    
                    for n in range(len(imgs)):
                        
                        nel=0
                        for el in imgtags:
                            if any([str(n+1)+".png" in el,str(n+1)+".jpg" in el,str(n+1)+".gif" in el]):
                                nel=nel+1
                                
                                dirst=self.cd+'Email/'+imgs[n]
                                with open(dirst.replace('/','\\'), 'rb') as fp:
                                    msgImage = fp.read()
                                    data_base64 = base64.b64encode(msgImage)
                                    data_base64 = data_base64.decode()
                                lsubs.append(data_base64)
                                break
                        if nel==0:
                            lsubs.append(self.emptyflag)

                    nsubstitutions=len(imgs)
                    isubstitutions=[kids for kids,ids in enumerate(imgs) if lsubs[kids]!=self.emptyflag]
                    lstrrep=[''.join(['"',"data:image/jpeg;base64,",str(el),'"']) if el!=self.emptyflag else self.emptyflag for el in lsubs]

                    string=Execution.nsubstitute(strbase,lstrrep,nsubstitutions,isubstitutions,a,b,encoding)

                    strbase=string
                    a="["
                    b="]"
                    nsubstitutions=len(Execution.extractall(strbase,{"[]":[[a,b]]})["[]"])
                    isubstitutions=[kids for kids,ids in enumerate(Execution.extractall(strbase,{"[]":[[a,b]]})["[]"]) if all([ids=="","if" not in ids])]
                    lstrrep=[lemailsubstitutions[isubstitutions.index(num)+2]  if num in isubstitutions else self.emptyflag for num in range(nsubstitutions)]

                    string2=Execution.marker_nsubstitute(strbase,lstrrep,nsubstitutions,isubstitutions,a,b)

                    msg.attach(MIMEText(string2, 'html'))

                    server = smtplib.SMTP(self.nserver,587)
                    server.starttls()
                    server.ehlo()
                    try:
                        server.login(self.fromaddr, self.passw)
                    except:
                        messagebox.showerror("Error", "Error en email o contraseña.")
                    text= msg.as_string()
                    server.sendmail(self.fromaddr, lemailsubstitutions[0].split(','), text)
                    server.quit()

            dictv={"Email":args[0],"Password":"Rellenar","Server":args[2]}
            with open(self.cd+"Config.json","w") as f:
                json.dump(dictv,f)
    #            shutil.rmtree(self.cd+"\\Email\\Email_archivos")
    #            os.remove(self.cd+"\\Email\\Email.htm")
    #            os.chdir(self.cd+"\\Enviados")
    #            flst=[el for el in os.listdir() if el.endswith(".htm")]
    #            os.chdir(self.cd)
    #            shutil.move(self.cd+"\\Email\\Email.htm", self.cd+"\\Enviados\\Email"+str(len(flst)+1)+".htm")
    #            shutil.copytree(self.cd+"\\Email_archivos", self.cd+"\\Enviados\\Email_archivos"+str(len(flst)+1)) 
        except:
            messagebox.showerror("Error", "Verifique que los nombres originales de todos los archivos y carpetas del programa no han sido cambiados. Guarde la próxima vez el archivo en formato .docx en el mismo directorio que el .htm (PMEmails\Email) para poder añadirle a la carpeta de Enviados. que el número de columnas del archivo .xlsx coincide con el número de parejas de corchetes [] en el archivo .html.")
      
try:
    from bs4 import BeautifulSoup
    from docx import Document
except:
    messagebox.showerror("Error", "Para que funcione el programa debe estar instalado Python 3.7 al menos, además del paquete BeautifulSoup y docx.")

if __name__=="__main__":
    app=TkWindow()