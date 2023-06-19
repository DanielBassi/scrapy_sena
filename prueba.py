from tkinter import *
from tkinter import ttk
from tkinter import messagebox

from tkinter import scrolledtext as st
import sys
from tkinter import filedialog as fd
from tkinter import messagebox as mb

#Importar entorno de desarrollo de pruebas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotVisibleException

from selenium.common.exceptions import NoSuchElementException

#Correo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64

#Comparar archivos
import filecmp

#Importar Modulo PLatform para obtener informaci�n del ordenador
import platform

#Funci�n para obtener el nu�mero de sigundos trancurridos en el sistema
import time
import random
import re
from datetime import date
import datetime
import pytz

import tkinter as tk

#Copiar archivos
import shutil

class Interface:    
    
    def __init__(self):
        #Propiedades
        self.email_username = "pythonacountexample@gmail.com"
        self.email_password = "dljgcxhoxemozrwe"
        self.email_to = ['Luis Avila <zluisigloxx@gmail.com>']
        self.email_today = datetime.datetime.now((pytz.timezone('America/Bogota')))
        self.email_date_time = self.email_today.strftime("%d/%m/%Y, %H:%M:%S")
        
        self.original = "DB/REPORTE/REPORTE.txt"
        self.copia = "DB/REPORTE/REPORTE2.txt"

    def enviar_correo(self, descripcion=False):
                self.email_hora = datetime.datetime.now((pytz.timezone('America/Bogota'))).strftime("%H")
                message = MIMEMultipart()
                #saludo = 'HOLA' if lang=='es' else 'HI'
                message.attach(MIMEText( "Reporte de tu buen amigo el BOT" ))
                message['Subject'] = "Reporte calificaciones pendientes - Sena - " + datetime.datetime.now((pytz.timezone('America/Bogota'))).strftime("%d/%m/%Y, %H:%M:%S") if descripcion == False else descripcion
                message['From'] = self.email_username
                message['To'] = ','.join( self.email_to )
                
                part = MIMEBase('application', "octet-stream")
                part.set_payload(open(self.original, "rb").read())
                encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="REPORTE.txt"')
                message.attach(part)
                #Intentar enviar el correo
                stmp = smtplib.SMTP_SSL('smtp.gmail.com')
                self.email_hora = int(self.email_hora)
                print("hola")
                
                try:
                    stmp.login(self.email_username, self.email_password)
                    stmp.sendmail(self.email_username, self.email_to, message.as_string())
                    print("Correo enviado exitosamente.")
                except Exception as inst:
                    print("Ha ocurrido un error al intentar enviar correos electronicos")
                    print("OS error: {0}".format(inst))
                
                stmp.quit()

p=Interface()
p.enviar_correo()