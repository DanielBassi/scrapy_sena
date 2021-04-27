from tkinter import *
from tkinter import ttk

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

from tkinter import messagebox

#Importar Modulo PLatform para obtener información del ordenador
import platform

#Función para obtener el nuúmero de sigundos trancurridos en el sistema
import time
import random
import re
from datetime import date

import tkinter as tk
class Interface:
    def __init__(self):
        #Creación y configuracion de la ventana 
        self.window=Tk()
        self.window.title("BOT - Calificaciones SENA automatizadas")
        self.mainframe = ttk.Frame(self.window, padding="3 3 12 12")
        self.mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.window.resizable(0,0)

        #Obtener valores de User y Password
        user_pass = open('DB/USER.txt')
        self.user_pass = user_pass.read().split('|')
        
        #Columna Izquierda
        #Input Usuario
        ttk.Label(self.mainframe, text="Usuario").grid(column=1, row=2, sticky=(W, E))
        self.user = StringVar()
        entryUser = ttk.Entry(self.mainframe, width=20, textvariable=self.user)
        entryUser.grid(column=1, row=3, sticky=(W, E))
        entryUser.delete(0,END)
        entryUser.insert(0,self.user_pass[0])
        entryUser.focus()

        #Input Password
        ttk.Label(self.mainframe, text="Contraseña").grid(column=1, row=4, sticky=(W, E))
        self.passw = StringVar()
        entryPass = ttk.Entry(self.mainframe, width=20, textvariable=self.passw)
        entryPass.grid(column=1, row=5, sticky=(W, E))
        entryPass.insert(0, self.user_pass[1])

        #Checkbox Update Data Base
        self.update_DB = IntVar()
        checkbox = ttk.Checkbutton(self.mainframe, text="Actualizar DB", variable=self.update_DB).grid(column=1, row=6, sticky=(W, E))
    
        #Columna Derecha
        #Select Ficha
        ttk.Label(self.mainframe, text="Ficha").grid(column=2, row=2, sticky=(W, E))
        self.ficha = StringVar()
        self.chosen_ficha = ttk.Combobox(self.mainframe, width=110, state='readonly', textvariable=self.ficha)
        self.chosen_ficha.grid(column=2, row=3, sticky=(W, E))
        self.fichas = self.getDataBase('DB/FICHAS/FICHAS.txt')
        self.chosen_ficha['values'] = tuple( self.fichas )
        self.chosen_ficha.current()
        self.chosen_ficha.bind("<<ComboboxSelected>>", self.cargarEvidencia)

        #Select Evidencia
        ttk.Label(self.mainframe, text="Evidencia").grid(column=2, row=4, sticky=(W, E))
        self.evidencia = StringVar()
        self.chosen_evidencia = ttk.Combobox(self.mainframe, width=65, state='readonly', textvariable=self.evidencia)
        self.chosen_evidencia.grid(column=2, row=5, sticky=(W, E))

        #File browser Comentarios
        ttk.Label(self.mainframe, text="Archivo de comentarios").grid(column=2, row=6, sticky=(W, E))
        self.ruta_comentario = open('DB/COMENTARIOS/RUTA.txt').read()
        self.entryFileUpload = ttk.Entry(self.mainframe)
        self.entryFileUpload.grid(column=2, row=7, sticky=(W, E))
        self.entryFileUpload.insert(0, self.ruta_comentario)

        #Button action
        ttk.Button(self.mainframe, text="Iniciar", command=self.automatizacion).grid(column=1, row=8, sticky=(W, E))
        ttk.Button(self.mainframe, text="Observador", command=self.observador).grid(column=1, row=9, sticky=(W, E))      
        ttk.Button(self.mainframe, text="Salir", command=self.window.destroy).grid(column=1, row=10, sticky=(W, E))      
        self.window.mainloop()
    
    def cargarEvidencia(self, event):
        self.ficha_code = event.widget.get().split('__', 1)[0]
        self.evidencias = []
        try:
            self.evidencias = self.getDataBase('DB/EVIDENCIAS/'+self.ficha_code+'.txt')
            self.chosen_evidencia['values'] = tuple( self.evidencias )
        except:    
            self.chosen_evidencia['values'] = tuple( [] )
            self.chosen_evidencia.set('')
            messagebox.showinfo(message="No existen evidencias para la FICHA "+self.ficha.get(), title="Error del sistema")
    
    def auth(self):
        #Configurar Platform
        plataforma = platform.system()
        path_to_chromedriver = "EXE\chromedriver.exe"

        self.browser = webdriver.Chrome(executable_path = path_to_chromedriver)

        #URL que se quiere abrir
        url = 'https://sena.territorio.la/index.php?login=true'
        self.browser.get(url)

        #Asignar los valores al formulario de inicio de sesión
        self.browser.find_element_by_name('document').send_keys( self.user.get() )
        self.browser.find_element_by_name('password').send_keys( self.passw.get() )

        #Clic al botón de INGRESO
        self.browser.find_element_by_css_selector('.boton_verde').click()
    
    def observador(self):
        #Autentificar al usuario
        self.auth()
        
        for ficha in self.fichas:
            array_ficha = ficha.split('__')
            print(array_ficha[1])
            try:
                print("A ver que hace esto")
                evidencia_actual = self.getDataBase( 'DB/EVIDENCIAS/'+array_ficha[0]+'.txt' )
                
                #Redireccionar al listado de evidencias pertenecientes a la ficha
                #self.dirigir_a_evidencias( 'https://sena.territorio.la/perfil.php?id='+ array_ficha[0])
                time.sleep(2)
                self.browser.get( 'https://sena.territorio.la/perfil.php?id='+ array_ficha[0] )
                print('URL____________')
                #Dar clic en Evidencias
                salida = None
                print('INTENTO')
                while not salida:
                    try:
                        
                        self.browser.find_element_by_id('aTareas').click()
                        salida = True
                    except:
                        pass
                print('Me voy a '+array_ficha[1])
                
                for evidencia in evidencia_actual:
                    e = evidencia.split('__')
                    print('Comienzo')
                    #Dar clic en la Evidencia que se desea comenzara a calificar
                    salida = None
                    while not salida:
                        try:
                            self.browser.find_element_by_xpath("//a[contains(@onclick,'"+e[0]+"')]").click()
                            salida = True
                        except:
                            salida = None
                    
                    #Get all elements of father div #formCalificar
                    list_tables = None
                    while not list_tables:
                        try:
                            list_tables = self.browser.find_elements_by_css_selector("#formCalificar > table")
                        except NoSuchElementException:
                            list_tables = None
                    
                    #Cantidad de Registros calificados
                    count_qualifity = 0

                    print('Continuo')
                    for table in list_tables:
                        table_id = table.get_attribute('id').split('table-respuesta-', 1)[1]
                        print('Valido')
                        intentos = None
                        while not intentos:
                            try:
                                intentos = table.find_element_by_xpath("//table[@id='table-respuesta-"+table_id+"']/tbody/tr/td[@width='30%']/p").text.split(":")[-1].strip()
                            except NoSuchElementException:
                                intentos = table.find_element_by_xpath("//table[@id='table-respuesta-"+table_id+"']/tbody/tr/td[@width='30%']/span/p").text.split(":")[-1].strip()
                        
                        nota = table.find_element_by_xpath("//table[@id='table-respuesta-"+table_id+"']/tbody/tr/td[@width='6%']/input[@type='text']")

                        if nota.get_attribute('value') == 'Sn' and int(intentos) > 0:
                            count_qualifity = count_qualifity+1       
                            estudiante = table.find_element_by_xpath("//table[@id='table-respuesta-"+table_id+"']/tbody/tr/td[@width='59%']").text
                            
                            reporte = open('DB/REPORTE/'+array_ficha[1].replace('\n', '')+'/'+e[1].replace('\n', '')+'.txt', "w")
                            reporte.write( estudiante+'\n' )
                            reporte.close()
                        
                        reporte = open('DB/REPORTE/REPORTE.txt', "w")
                        reporte.write( array_ficha[1].replace('\n', '')+'|'+e[1].replace('\n', '')+'|'+str(count_qualifity)+'\n' )
                        reporte.close()
                    
                    #Redireccionar al listado de evidencias pertenecientes a la ficha
                    self.dirigir_a_evidencias( 'https://sena.territorio.la/perfil.php?id='+ array_ficha[0])
                
            except FileNotFoundError:
                messagebox.showerror(message="El archivo DB/EVIDENCIAS/"+array_ficha[0]+".txt de la FICHA "+array_ficha[1]+" no existe", title="Error del sistema")
        
    def automatizacion(self):
        
        #Autentificar al usuario
        self.auth()
        
        browser = self.browser    
        
        #Obtener el listado de FICHAS y sus respectivos nombres e ID's
        list_ficha_ids = None
        while not list_ficha_ids:
            try:
                list_ficha_ids = browser.find_elements_by_css_selector('.letras1')
            except NoSuchElementException:
                list_ficha_ids = None
        
        #Actualizar FICHAS.txt
        if self.update_DB.get() == 1:
            DB_file = open('DB/FICHAS/FICHAS.txt', "w")
            for element in list_ficha_ids:
                text = element.get_attribute("text")
                href = element.get_attribute("href")
                href = href.split("=", 1)[1]
                DB_file.write( href+'|'+text+'\n' )
            DB_file.close()

        if self.fichas == []:
            messagebox.showerror(message="Ha ocurrido un problema con las FICHAS, por favor vuelva a intentarlo", title="Error del sistema")
            browser.quit()
        
        #Redireccionar al listado de evidencias pertenecientes a la ficha
        self.dirigir_a_evidencias( 'https://sena.territorio.la/perfil.php?id='+self.ficha_code )

        #Obtener la lista de las evidencias        
        list_data_select_evidencias = None
        while not list_data_select_evidencias:
            try:
                list_data_select_evidencias = browser.find_elements_by_css_selector('.tituloTarea:first-child')
            except NoSuchElementException:
                list_data_select_evidencias = None
                    
        #Guardar en la DB la lista
        if self.update_DB.get() == 1:
            DB_file = open( 'DB/EVIDENCIAS/'+self.ficha_code+'.txt', "w" )
            for element in list_data_select_evidencias:
                text = element.get_attribute("text")
                text = re.sub("\s+", " ", text.strip())
                code = element.get_attribute("onclick")
                code = str(code.split("=", 1)[1].split("'")[0])
                DB_file.write( code+'|'+text+'\n' )
            DB_file.close()
            #messagebox.showinfo(message="El archivo EVIDENCIAS/"+self.ficha_code+".txt ha sido actualizado con exito!", title="Actualización de Base de Datos")

        if self.evidencias == []:
            messagebox.showerror(message="Ha ocurrido un problema con las EVIDENCIAS, por favor vuelva a intentarlo", title="Error del sistema")
            browser.quit()

        #Dar clic en la Evidencia que se desea comenzara a calificar
        salida = None
        while not salida:
            try:
                browser.find_element_by_xpath("//a[contains(@onclick,'"+self.evidencia.get().split('__')[0]+"')]").click()
                salida = True
            except:
                salida = None

        #Get all elements of father div #formCalificar
        list_tables = None
        while not list_tables:
            try:
                list_tables = browser.find_elements_by_css_selector("#formCalificar > table")
            except NoSuchElementException:
                list_tables = None
        
        #Cantidad de Registros calificados
        count_qualifity = 0

        for table in list_tables:
            table_id = table.get_attribute('id').split('table-respuesta-', 1)[1]
            
            intentos = None
            while not intentos:
                try:
                    intentos = table.find_element_by_xpath("//table[@id='table-respuesta-"+table_id+"']/tbody/tr/td[@width='30%']/p").text.split(":")[-1].strip()
                except NoSuchElementException:
                    intentos = table.find_element_by_xpath("//table[@id='table-respuesta-"+table_id+"']/tbody/tr/td[@width='30%']/span/p").text.split(":")[-1].strip()
            
            nota = table.find_element_by_xpath("//table[@id='table-respuesta-"+table_id+"']/tbody/tr/td[@width='6%']/input[@type='text']")
    
            if nota.get_attribute('value') == 'Sn' and int(intentos) > 0:
                nota.clear()
                nota.send_keys(100)
                table.find_element_by_xpath("//table[@id='table-respuesta-"+table_id+"']/tbody/tr/td[@width='6%']/input[@type='checkbox']").click()
                table.find_element_by_xpath("//a[@href='javascript:verMasComentariosTareas("+table_id+");']").click()
                table.find_element_by_xpath("//a[contains(@onclick, 'ComentarioAUser_"+table_id+"')]").click()
                comentario = table.find_element_by_xpath("//*[@id='ComentarioAUser_"+table_id+"']")
                comentario.find_element_by_class_name("commentMark").send_keys( self.obtenerComentario() )
                count_qualifity = count_qualifity+1

        #messagebox.showinfo(message="Se han calificado: %i" %count_qualifity, title="Mensaje del sistema")
    
    def obtenerComentario(self):
        comentario = open(self.ruta_comentario, encoding='utf-8').read().split('\n')
        return comentario[ random.randint( 0, len(comentario) ) ]
        
    def dirigir_a_evidencias(self, url):
        time.sleep(1)
        self.browser.get( url )
        #Dar clic en Evidencias
        salida = None
        while not salida:
            try:
                self.browser.find_element_by_id('aTareas').click()
                salida = True
            except:
                pass
    
    def getDataBase(self, file, separator = '|'):
        arr = []
        with open(file) as inFile:
            arr = [line for line in inFile]
            arr = [(f.split( separator )[0]+'__'+f.split( separator )[1]) for f in arr]
        return arr
        
aplicacion1=Interface()