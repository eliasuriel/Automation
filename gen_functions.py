import pandas as pd
import numpy as np
import math
import string
import os
import re
import csv
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import NamedStyle
from openpyxl.worksheet.table import Table, TableStyleInfo 
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import easygui
import sys
from global_consts import *
from global_variables import *


def hoja_excel(file_path,nombre_hoja):
    file = pd.ExcelFile(file_path)
    if nombre_hoja in file.sheet_names:
        hoja = file.parse(nombre_hoja)
        return hoja
    else:
        easygui.msgbox("La hoja {} no se encontró en {}. \nSaliendo del script...".format(nombre_hoja, file_path), "ERROR")
        sys.exit(0)

def month_to_index(month: str) -> int:
    month_order = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}
    return month_order.get(month.upper(), 0)

# Function to sort projects by month
def sort_projects_by_month(projects):
    return sorted(projects, key=lambda x:(x[0],month_to_index(x[1]),x[2], x[3],x[4], x[5], x[6], x[7],x[8]))

def change_columns(hoja):
    nuevos_nombres = hoja.iloc[0]  # Obtén los nombres de la primera fila
    hoja.columns = nuevos_nombres 
    hoja = hoja.shift(-1)
    return hoja

def columnas(hoja,nombres_columna,index):
    hoja[nombres_columna[index]] = hoja[nombres_columna[index]].fillna(0)
    columna = hoja[nombres_columna[index]]
    return columna

def minusculas(columna):
    col_lower = [str(name).lower() for name in columna]
    return col_lower

def borrar_segunda_palabra(cadena):
    palabras = cadena.split()
    palabras.pop(1)
    nueva_palabra = ' '.join(palabras)
    return nueva_palabra

def existe_persona(lista_personas, nombre):
    for persona in lista_personas:
        if persona.nombre() == nombre:
            return True
    return False

def convertir_horas_a_float(arreglo):
 numeros_float = []
 for hora in arreglo:
     numero_str = hora.replace('h', '')  # Eliminar la letra 'h'
     numero_float = float(numero_str)  # Convertir a punto flotante
     numeros_float.append(numero_float)
 return numeros_float

def comun_prime_opx(personas_opx,col_prime):
    personas_opx_prime = []
    personas_prime = []
    for nombre_opx in personas_opx:
        for nombre_prime in col_prime:
            if (nombre_opx == '0') or ('other projects and commitments' in (nombre_prime)):
                continue
            elif ('egb' in nombre_prime):
                nuevo_elemento = borrar_segunda_palabra(nombre_prime)
                if (nombre_opx) in (nuevo_elemento):
                    personas_opx_prime.append(nombre_opx)
                    personas_prime.append(nombre_prime)
                       
    return personas_opx_prime,personas_prime

def personas_difname(personas_opx_prime,personas_opx,col_prime,prime_list):
    list_black = []
    prime_black_list = []
    for nombres in personas_opx:
        if nombres in personas_opx_prime:
            pass
        else:
            list_black.append(nombres)

    
    for nombre_prime in col_prime:
        if nombre_prime in prime_list or ('other projects and commitments' in (nombre_prime)) :
            continue
        else:
            if "bgsw" in nombre_prime:
                prime_black_list.append(nombre_prime)
            

    return list_black,prime_black_list


def personas_prime(prime_person,col_prime):
    for nombre in col_prime:
        if "bgsw" in nombre:
            prime_person.append(nombre)

    return prime_person

def agregar_personas_opx(opx_person,col_resource,black_list):
    
    for nombre_resource in col_resource:
        if nombre_resource in opx_person:
            continue
        elif nombre_resource in black_list:
            continue
        else:
            opx_person.append(nombre_resource)

    return opx_person

# def agregar_personas_opx(col_resource,black_list):    
#     opx_person = []
#     for nombre_resource in col_resource:
#         if nombre_resource in opx_person or nombre_resource in black_list:
#             continue
#         else:
#             opx_person.append(nombre_resource)
    
#     # for nombre in col_resource:
#     #     nombres_separados = nombre.split(',')
#     #     for n in nombres_separados:
#     #         if n.strip() not in black_list and n.strip() not in opx_person:
#     #             opx_person.append(n.strip())
#     return opx_person

def proyectos(col_proyectos):
    opx_proy = []
    for nom_proy in col_proyectos:
        if(len(opx_proy) == 0) and nom_proy != '0':
            opx_proy.append(nom_proy)
        elif nom_proy in str(opx_proy):
            continue
        else:
            if nom_proy != '0':
             opx_proy.append(nom_proy)
    return opx_proy

def horas_proyecto_opx(nombre_persona, nombre_proyecto,columna_proyecto,columna_persona,columna_horas):
    index = 0
    suma_horas = 0
    for proyecto in columna_proyecto:

        if (proyecto == nombre_proyecto or proyecto in nombre_proyecto) and (nombre_persona in columna_persona[index] or nombre_persona == columna_persona[index]):
            suma_horas += float(columna_horas[index])
        index += 1

    return suma_horas


        
    
  
def proyecto_persona(hoja, colum_res, colum_proj, nombre_persona):
    pers_proyect = []
    for index, row in hoja.iterrows():
        if nombre_persona in colum_res[index] and colum_proj[index] not in pers_proyect:
            pers_proyect.append(colum_proj[index])
              
    pers_proyect = minusculas(pers_proyect)
    return pers_proyect

def comparar_nombres(nombre1, nombre2):
    # Dividir los nombres en palabras individuales
    palabras_nombre1 = nombre1.split()
    palabras_nombre2 = nombre2.split()

    # Verificar la similitud
    similitud = len(set(palabras_nombre1) & set(palabras_nombre2)) / float(len(set(palabras_nombre1) | set(palabras_nombre2)))

    return similitud

def relacionar_palabras(arr1, arr2):
    relacion = {}
    for nombre1 in arr1:
        mejor_coincidencia = ""
        longitud_coincidencia = 0
        for nombre2 in arr2:
            lcs = comparar_nombres(nombre1, nombre2)
            if lcs == 0:
                pass
            else:
                if lcs > longitud_coincidencia:
                    mejor_coincidencia = nombre2
                    longitud_coincidencia = lcs
        relacion[nombre1] = mejor_coincidencia
    return relacion

  
def result_button(output):
    if output:
        #Continue
        pass 
    else:
        #Exit
        msg = easygui.msgbox("Finalizando Script...", "Salir")
        sys.exit(0)
        

def encontrar_numero_mas_alto(directorio):
    # Obtener la lista de archivos en el directorio
    archivos = os.listdir(directorio)

    max_numero = 0

    # Patron de los nombres de archivo de salida
    patron = re.compile(r'Output-(\d+)')

    # Iterar sobre los archivos y encontrar el número más alto
    for archivo in archivos:
        match = patron.match(archivo)
        if match:
            numero = int(match.group(1))
            if max_numero is None or numero > max_numero:
                max_numero = numero

    return max_numero


def busqueda_nombres(hoja1,hoja2,mes_prime,mes_opx,anio,column_titles_opx):
    hoja1_columna_nombre = "Resource Name"
    hoja2_columna_nombre = "Resource"
   
    for i, col_header in enumerate(column_titles_opx):
        if mes_opx in col_header and anio in col_header:
            hoja2_columna_horas = col_header
    # Iterar sobre las filas de la hoja2
    for index, row in hoja2.iterrows():
        nombre_hoja2 = row[hoja2_columna_nombre].split()
        find_string = False
    
        # Iterar sobre las filas de la hoja1
        for index2, row2 in hoja1.iterrows():
            hoja1_columna1 = row2[hoja1_columna_nombre]
        
            find_string = True
            for count, ele in enumerate(nombre_hoja2):
                ele = ele.replace(",", "")

                if hoja1_columna1.find(ele) == -1:
                    find_string = False
                    break
    
            # Si se encontró una coincidencia, actualizar las horas en la hoja2
            if find_string:
                contenido = row2[mes_prime]
                hoja2.at[index, hoja2_columna_horas] = contenido

    return hoja1,hoja2

def relacionar_nombres(arr1, arr2):
    relacion = {}
    for nombre1 in arr1:
        mejor_coincidencia = ""
        longitud_coincidencia = 0
        for nombre2 in arr2:
            lcs = longest_common_substring(nombre1, nombre2)
            if len(lcs) > longitud_coincidencia:
                mejor_coincidencia = nombre2
                longitud_coincidencia = len(lcs)
        relacion[nombre1] = mejor_coincidencia
    return relacion

def longest_common_substring(s1, s2):
    m = len(s1)
    n = len(s2)
    matriz = [[0] * (n + 1) for _ in range(m + 1)]
    longitud_max = 0
    final_indice = 0
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if s1[i - 1] == s2[j - 1]:
                matriz[i][j] = matriz[i - 1][j - 1] + 1
                if matriz[i][j] > longitud_max:
                    longitud_max = matriz[i][j]
                    final_indice = i
            else:
                matriz[i][j] = 0
    return s1[final_indice - longitud_max:final_indice]

# Función para convertir todo a mayúsculas
def convertir_mayusculas(texto):
 return texto.upper()

# Función para capitalizar cada palabra
def capitalizar_palabras(texto):
 return ' '.join(word.capitalize() for word in texto.split())



def validate_user(username, password):
    for user in user_credentials:
        if user[0] == username and user[1] == password:
            return True
    return False


def login(entry_username, entry_password, window):
    search_username = entry_username.get()
    search_password = entry_password.get()
    global logged

    print("search_username:", search_username)
    print("search_password:", search_password)
        
    if validate_user(search_username, search_password) == True:
        logged = True
        window.quit()
        window.destroy()
    else:
        messagebox.showerror("Error", "Invalid username or password. Please try again.")


def login_window():
    # Create the main window
    window = tk.Tk()
    window.title("Login")
    window.geometry("300x150")  # Set the width and height of the window

    # Create and position the username label and entry
    label_username = tk.Label(window, text="Username:")
    label_username.pack()
    entry_username = tk.Entry(window)
    entry_username.pack()

    # Create and position the password label and entry
    label_password = tk.Label(window, text="Password:")
    label_password.pack()
    entry_password = tk.Entry(window, show="*")
    entry_password.pack()

    # Create and position the login button
    button_login = tk.Button(window, text="Login", command=lambda: login(entry_username, entry_password, window))
    button_login.pack()
    window.mainloop()
    
    if logged == False:
        sys.exit(0)

def ord_alfabet(arreglo):
 arreglo_ordenado = sorted(arreglo)
 return arreglo_ordenado



