# -*- coding: utf-8 -*-
"""
Created on Sat Feb 19 14:04:39 2022

@author: Jaime García Chaparro
"""

import pandas as pd
import openpyxl
import os
import datetime

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

email = '' # Correo desde el que se envía
password = '' # Contraseña del correo desde el que se envía
send_to_email = '' # Correo al que se envía

# Variables básicas

script_dir = os.path.dirname(__file__)
num_semana = datetime.datetime.now().isocalendar()[1]
nombre_menu  =  f'Menu semana {num_semana}.xlsx'
nombre_plantilla = 'Plantilla.xlsx'

comidas = ['desayuno', 'almuerzo', 'cena']
dias = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo']

#%% Importar y procesar tablas

# Criterios
crit = pd.read_excel(os.path.join(script_dir, 'Criterios.xlsx'), index_col= 'comida')

## Convertir en criterios en listas
for i in range(0, 3):
    for j in range(0, 7):
        crit.iloc[i, j] = crit.iloc[i, j].split('; ')

# Recetas
rec = pd.read_excel(os.path.join(script_dir, 'Recetas.xlsx'))
#rec.set_index('plato', inplace = True)

## Convertir en lista los ingredientes
for i in range(len(rec)):
    ing_part = rec.loc[i, 'ingredientes'].split('; ')
    ing_dic = []
    for el in ing_part:
        spl = el.split(', ')
        rec_dic = {spl[1] : int(spl[0])}
        ing_dic.append(rec_dic)
        
    rec.loc[i, 'ingredientes'] = ing_dic

# Lista de la compra
lcom = pd.DataFrame(data = None, columns = ['producto', 'cantidad', 'seccion'])
lcom.set_index('producto', inplace = True)

# Productos
prods = pd.read_excel(os.path.join(script_dir, 'Productos.xlsx'), index_col= 'producto')

# Cuadro final
menu = pd.DataFrame(data = None, index = comidas, columns = dias)

#%% Definir funciones

def anadir_comida(criterio, comida, dia):
    """Añade una comida al cuadro final."""
    
    # Selecciona comida con criterio al azar
    row = rec.loc[rec.categoria == criterio].sample()
    
    # Añade la comida al menú
    elaboracion = row['plato'].iloc[0]
    try:
        menu.loc[comida, dia] += f', {elaboracion}'
    except:
        menu.loc[comida, dia] = f'{elaboracion}'
    
    # Añadir ingredientes a lista de la compra
    ings = row['ingredientes']
    for ing in ings.iloc[0]:
        for prod, q in ing.items():
            anadir_compra(prod, q)
    
def anadir_compra(elemento, cantidad):
    """Añade un ingrediente a la lista de la compra."""
    
    try:
        lcom.loc[elemento, 'cantidad'] += cantidad
    except:
        lcom.loc[elemento, 'cantidad'] = cantidad
        try:
            lcom.loc[elemento, 'seccion'] = prods.loc[elemento, 'seccion']
        except:
            lcom.loc[elemento, 'seccion'] = 'Por determinar'

#%% Función para rellenar el menú

def crear_menu():
    for c in comidas:
        for d in dias:
            cr = crit.loc[c, d] # Criterios de la comida
            for crt in cr: # Por cada elemento en criterios...
                anadir_comida(crt, c, d)

crear_menu()

#%% Exportar menú

plantilla = openpyxl.load_workbook(os.path.join(script_dir, nombre_plantilla))
ex_menu = plantilla['Menú']
ex_lcom = plantilla['Lista']

# Rellenar menú en excel
l_pl = ['C', 'D', 'E', 'F', 'G', 'H', 'I']
n_pl = [4, 5, 6]
n_pl_str = [str(n) for n in n_pl]

for l, d in zip(l_pl, dias):
    for n, c in zip(n_pl_str, comidas):
        ex_menu[f'{l}{n}'] = menu.loc[c, d]
        
# Rellenar lista de la compra
# Usar lista con índice reseteado por comodidad
lcom.sort_values(by = 'seccion', inplace = True)
lcom_rst = lcom.copy()
lcom_rst.reset_index(inplace = True)

n_inicio = 2 # Fila en la que empieza la lista de la compra

for i in range(0, len(lcom_rst)):
        ex_lcom[f'A{str(i + n_inicio)}'] = lcom_rst.iloc[i, 0].capitalize()
        ex_lcom[f'B{str(i + n_inicio)}'] = lcom_rst.iloc[i, 1]
        ex_lcom[f'C{str(i + n_inicio)}'] = lcom_rst.iloc[i, 2].capitalize()
        
# Guardar el menú
plantilla.save(os.path.join(script_dir, nombre_menu))

#%% Enviar correo

def enviar_correo():
    """Envía un correo con el menú y la lista de la compra."""

    print('Enviando correo...')
        
    email_user = email
    email_password = password
    email_send = send_to_email
    subject = f'Menú semana {num_semana}'
    
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subject
    
    body = f'Enviando menú de la semana {num_semana}'
    msg.attach(MIMEText(body,'plain'))
    
    filename = nombre_menu
    attachment = open(nombre_menu,'rb')
    
    part = MIMEBase('application','octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',"attachment; filename= "+ filename)
    
    msg.attach(part)
    text = msg.as_string()
    serve = smtplib.SMTP('smtp.gmail.com',587)
    serve.starttls()
    serve.login(email_user,email_password)
    
    serve.sendmail(email_user,email_send,text)
    serve.quit()
    attachment.close()
    
    print('Correo enviado.')

enviar_correo()