from gen_functions import *
from persona import *

#############################################
#        INICIALIZAR INPUTS
############################################
#Ventana de login
login_window()

# Mostrar un cuadro de mensaje de bienvenida
choice_index = easygui.indexbox(msg, title, initial_choices)

if choice_index == DEFAULT_EXIT:
    msg = easygui.msgbox("Finalizando Script...", "Salir")
    sys.exit(0)
elif choice_index == DEFAULT_CONTINUE:
    while True:
        file_path = filedialog.askopenfilename(title="Escoge el archivo PRIME", filetypes=(("Archivo de Excel", "*.xlsx"),))
        if file_path:
            break
        result_button(easygui.ccbox(msg,title,choices))

    # Pedir al usuario que elija el archivo de Excel con datos de OPX
    while True:
        result_button(easygui.ccbox("Elija su archivo de Excel con datos de OPX","",choices))
        file_path_opx = filedialog.askopenfilename(title="Escoge el archivo OPX", filetypes=(("Archivo de Excel", "*.xlsx"),))
        if file_path_opx:
            break
else:
    current_path = os.getcwd()
    current_path = os.path.dirname(current_path)
    default_settings = True
    file_path = current_path + PRIME_PATH
    file_path_opx = current_path + OPX_PATH
    anio = DEFAULT_YEAR
    hours_per_day = DEFAULT_WK_HOURS
    directorio = current_path + DEFAULT_OUT_PATH

# Crear objetos de archivo Excel a partir de los archivos seleccionados
df_prime =hoja_excel(file_path,hoja_prime)
df = hoja_excel(file_path_opx,hoja_opx)
df = change_columns(df)
nombre_columna_p = df_prime.columns
column_titles = df_prime.columns.tolist()
column_titles_opx = df.columns.tolist()
nombres_columnas = df.columns

col_nombres_opx = columnas(df,nombres_columnas,8)
col_resource_opx = columnas(df,nombres_columnas,7)
col_leader_opx = columnas(df,nombres_columnas,4)
col_proyecto = columnas(df,nombres_columnas,0)
col_nombres_prime = columnas(df_prime,nombre_columna_p,0)
col_description = columnas(df,nombres_columnas,3)

nombres_opx_persona = []
nombres_prime_persona = []
index = 0

# Convertir a minúsculas los nombres de columnas
#col_nombres_opx = minusculas(col_nombres_opx)
col_nombres_prime = minusculas(col_nombres_prime) #nombres y proyectos de prime
col_nombres_opx = minusculas(col_nombres_opx)
col_resource_opx = minusculas(col_resource_opx) #nombres de opx
black_list = minusculas(black_list)
col_leader_opx = minusculas(col_leader_opx)
col_proyecto = minusculas(col_proyecto)
col_description = minusculas(col_description)

#Nombres de personas en opx
nombres_opx_persona = agregar_personas_opx(nombres_opx_persona,col_resource_opx,black_list)
nombres_opx_persona= minusculas(nombres_opx_persona)

#Nombres de personas en prime
nombres_prime_persona = personas_prime(nombres_prime_persona,col_nombres_prime)
nombres_prime_persona = minusculas(nombres_prime_persona)

personas_opx_prime = []
prime_list = []
list_black = []
prime_list_black = []
#Personas en comun opx_prime
personas_opx_prime, prime_list = comun_prime_opx(nombres_opx_persona,nombres_prime_persona)

list_black, prime_list_black  = personas_difname(personas_opx_prime,nombres_opx_persona,col_nombres_prime,prime_list)

####COMPARAR PERSONAS DE OPX Y PRIME####
similitudes = relacionar_palabras(nombres_opx_persona,nombres_prime_persona)
similitud_lista  = []
for similitud1,similitud2 in similitudes.items(): 
    if len(similitud2) == 0:
        pass
    else:
        #print(f"{similitud1} -> {similitud2}")
        similitud_lista.append(similitud1)

################################
#   SELECCION DE PARAMETROS
##################################

meses = []
proyectos_opx = []
personas_display = []

similitud_lista = ord_alfabet(similitud_lista)
for elemento in similitud_lista:
    personas_display.append(capitalizar_palabras(elemento))

#Se le pide al usuario que seleccione las personas
while  True:
    seleccion_persona = easygui.multchoicebox(f"Seleccione los asociados a analizar","Asociados",personas_display)
    if seleccion_persona:
        break
    else:
        msg = easygui.msgbox("Finalizando Script...", "Salir")
        sys.exit(0)

seleccion_persona = minusculas(seleccion_persona)

#Se guardan los titulos de las columnas de meses de prime
for nombre in column_titles:
    if nombre == FIRST_COLUMN_NAME_PRIME:
        continue
    else:
        meses.append(nombre)

#El usuario selecciona los meses
while True:
    seleccion_meses= easygui.multchoicebox(f"Selecciona en qué meses deseas trabajar? ","Meses",meses)
    if seleccion_meses:
        nummes = len(seleccion_meses)
        break
    else:
        msg = easygui.msgbox("Finalizando Script...", "Salir")
        sys.exit(0)


#El usuario elige los proyectos
proy_person = []
for persona in seleccion_persona:
    proy_person = proy_person + proyecto_persona(df,col_resource_opx,col_description,persona)
proy_person = list(set(proy_person))
proy_person = ord_alfabet(proy_person)
projects_display = ord_alfabet(projects_display)

if len(projects_display) > 1:
    while True:        
        seleccion_proyectos= easygui.multchoicebox(f"Selecciona en qué proyectos deseas trabajar? ","Proyectos",projects_display)
        if seleccion_proyectos:
            break
        else:
            msg = easygui.msgbox("Finalizando Script...", "Salir")
            sys.exit(0)
else:
    result_button(easygui.ccbox("El único proyecto encontrado fue: {}".format(proy_person[0]),"Proyecto",choices)) 
    seleccion_proyectos = proy_person

seleccion_proyectos = minusculas(seleccion_proyectos)
projects_display = minusculas(projects_display)

if default_settings == False:
    #Se ingresa el numero de horas laborales
    while True:
        hours_per_day = float(easygui.enterbox("¿Cuántas horas laborales tiene un día? (8, 8.25, 9...)"))
        if hours_per_day:
            if hours_per_day < 1:
                easygui.msgbox("Las horas laborales deben ser mayor a 1","ERROR")
            else:
                break
        else:
            sys.exit(0)

    yearFound = True

    while yearFound:
        anio = str(easygui.enterbox("En que año quieres trabajar? "))
        for titulo in column_titles_opx:
            if str(anio) in titulo:
                yearFound = False
                break
        else:
            easygui.msgbox("El año {} no fue encontrado en los títulos de las columnas del archivo OPX.\nFinalizando script...".format(anio),"ERROR")
            sys.exit(0) 
        

#Nombres de proyectos en opx
proyectos_opx = proyectos(col_proyecto)


#############################
# Crear Archivo Output
############################

if default_settings == False:
    while True:
        result_button(easygui.ccbox("Ingresa la carpeta donde deseas guardar el archivo de salida",choices))
        directorio = easygui.diropenbox()
        if directorio:
            break

default_settings = False        

max_numero = encontrar_numero_mas_alto(directorio)
archivo = OUTPUT_FILE_PREFIX + str(max_numero+1) + ".xlsx"
# Crear la ruta completa para el archivo de salida
ruta_completa = os.path.join(directorio, archivo)

# Verificar si el archivo de salida ya existe
if not os.path.exists(ruta_completa):
    # Si no existe, crear un nuevo libro de Excel y hojas
    workbook = openpyxl.Workbook()      
    hoja2 = workbook.active  
    hoja2.title = SHEET_2_NAME + str(anio)            
          
    hoja = workbook.create_sheet(title=str(anio)) 
    hoja.title = str(anio)               
           
    workbook.save(ruta_completa)        
    easygui.msgbox(f"El archivo {archivo} fue creado en: {directorio}")
else:
    # Si ya existe, cargar el libro existente y crear hojas
    easygui.msgbox(f"El archivo {archivo} ya existe en: {directorio}.")
    workbook = openpyxl.load_workbook(ruta_completa)
    hoja2 = workbook.create_sheet             
    hoja2.title = SHEET_2_NAME + str(anio)
    hoja = workbook.create_sheet              
    hoja.title = str(anio)               

    workbook.save(ruta_completa)


#####################
# Cell Colors and Styles
####################

# Definir estilos de celdas
relleno_titulos = PatternFill(start_color="A9C4E7", end_color="A9C4E7", fill_type="solid")
relleno_verde = PatternFill(start_color="7ce12e", end_color="7ce12e", fill_type="solid")
relleno_rojo = PatternFill(start_color="f25757", end_color="f25757", fill_type="solid")
relleno_naranja = PatternFill(start_color="feb055", end_color="feb055", fill_type="solid")
negritas = Font(bold=True)

########################
# Imprimir en output
#######################

celda = hoja["B1"]
celda.value = "Hours/day"
celda.fill = relleno_titulos
celda.font = negritas

celda = hoja["C1"]
celda.value = hours_per_day
celda.fill = relleno_titulos
celda.font = negritas

# Inicializar variables para filas y columnas de la segunda hoja, la de asociados
num_columna_2 = TABLE_START_COLUMN
num_fila_2 = TABLE_START_ROW

# Agregar títulos a las celdas de la hoja de asociados
for i in range(len(OUTPUT_HEADERS)):
    celda = hoja2.cell(row=num_fila_2, column=num_columna_2 + i, value=OUTPUT_HEADERS[i])
    celda.fill = relleno_titulos

num_fila_2 += 1

seleccion_proyectos = ord_alfabet(seleccion_proyectos)
seleccion_persona = ord_alfabet(seleccion_persona)
componentes = ord_alfabet(componentes)
componentes = minusculas(componentes)

cont = 0
Personas = []



#for gen_project in projects_display:
for proyecto in seleccion_proyectos:
    for mes in seleccion_meses:
        persona_opx = ""
        proyecto_opx = ""
        project_found_opx = ""
        nombre = ""
        opx_component = ""
        proyecto_prime = "N/A"
        horas_opx = 0
        suma_opx = 0
        horas_prime = 0
        suma_prime = 0
        relation = 0

        col_mes_prime = df_prime[mes]
        col_mes_prime = col_mes_prime.fillna('0')
        col_mes_prime = convertir_horas_a_float(col_mes_prime)

        for i, col_header in enumerate(column_titles_opx):
            if mes in col_header and anio in col_header:
                col_mes = columnas(df,nombres_columnas,i)

        for index, row in df.iterrows():
            persona_opx = col_resource_opx[index]
            proyecto_opx = col_description[index]
            horas_opx = float(col_mes[index]*hours_per_day)

            if (existe_persona(Personas,persona_opx)):
                    pass
            else:
                if persona_opx in seleccion_persona:
                    per = Persona(persona_opx)
                    Personas.append(per)
                else:
                    pass

            if proyecto_opx == '0':
                continue
            elif (proyecto in proyecto_opx) and (persona_opx in seleccion_persona):
                indice = index -1
                while col_resource_opx[indice] in nombres_opx_persona:
                    indice -= 1
                opx_component = col_resource_opx[indice]


                for componente in componentes:
                    if componente in opx_component:
                        break
                else:
                    componente = "N/A" 
            
                nombre = similitudes[persona_opx] #coincidencia en opx y prime

                conditional = False
                
                for index2, row2 in df_prime.iterrows():
                    #print(f'nombre: {nombre}\tprime: {col_nombres_prime[index2]}')
                    if nombre in col_nombres_prime[index2]:
                        person_prime = col_nombres_prime[index2]
                        conditional = True
                        continue
                    elif (col_nombres_prime[index2] in nombres_prime_persona):
                        if conditional:
                            break
                        else:
                            continue
                    elif (conditional) and (proyecto in col_nombres_prime[index2]):# and (componente in col_nombres_prime[index2]):
                        proyecto_prime = col_nombres_prime[index2]
                        horas_prime += col_mes_prime[index2]          

                if horas_opx:
                    relation = float(horas_prime/horas_opx)

                #cont += 1

                for name in Personas:
                    #print("nombre actual", name._nombre, "\n")
                    #print(proyecto_opx,proyecto_prime)
                    nombre = similitudes[name._nombre]
                    if persona_opx in name._nombre:
                        if proyecto in proyecto_opx and proyecto_opx != '0':
                            if componente in opx_component:
                                #print(name._nombre)
                                #print("proyecto",proyecto_opx,"componente",*componente,"horas",horas_opx, "mes", mes)
                                name.agregar_proyecto_simple(proyecto_opx,mes,"opx")
                                name.agregar_componente_a_proyecto(proyecto_opx,componente,horas_opx,mes)
                    if nombre in person_prime:
                        if proyecto in proyecto_prime:
                            name.agregar_proyecto_simple(proyecto_prime,mes,"prime")
                            name.agregar_horas_a_componente(proyecto_prime,componente,horas_prime,mes)
                        
                    
                            

                
                persona_opx = ""
                proyecto_opx = ""
                project_found_opx = ""
                nombre = ""
                opx_component = ""
                proyecto_prime = "N/A"
                horas_opx = 0
                suma_opx = 0
                horas_prime = 0
                suma_prime = 0
                relation = 0
   
comp = ""
contador = 0
unique_lines = set()  # Initialize a set to keep track of unique lines

for personitas in Personas:
    for proyecto in personitas.get_nombres_proyectos():
            for mes in seleccion_meses:
                for componente in personitas.get_componentes(proyecto,mes):
                    comp = componente['nombre_componente']
                    for project_default in projects_display:
                        for proyecto_prime in personitas.get_nombres_proyectos_por_fuente("prime"):
                            if (len(personitas.get_nombres_proyectos_por_fuente("opx")) > 0):
                                for proyecto_opx in personitas.get_nombres_proyectos_por_fuente("opx"):
                                    if project_default in proyecto_prime and project_default in proyecto_opx:
                                        if comp in componentes:
                                            horas_opx = personitas.get_horas_componente(proyecto_opx, comp, mes)
                                            horas_prime = personitas.get_horas_componente(proyecto_prime, comp, mes)
                                        else:
                                            horas_opx = personitas.get_horas_componente(proyecto_opx, "N/A", mes)
                                            horas_prime = personitas.get_horas_componente(proyecto_prime, "N/A", mes)

                                        if horas_opx:
                                            relation = float(horas_prime / horas_opx)
                                        else:
                                            relation = 0
                                        
                                        #print(project_default, mes, comp, personitas._nombre, proyecto_opx, proyecto_prime, horas_opx, horas_prime, relation)
                                        line = (project_default, mes, comp, personitas._nombre, proyecto_opx, proyecto_prime, horas_opx, horas_prime, relation)
                                        unique_lines.add(line)  # Add the line to the set
                                        
                                        # print(contador)
                                        contador = contador + 1
                                        break
                            else:
                                if project_default in proyecto_prime:
                                    if comp in componentes:
                                        horas_opx = 0
                                        horas_prime = personitas.get_horas_componente(proyecto_prime, comp, mes)
                                    else:
                                        horas_opx = 0
                                        horas_prime = personitas.get_horas_componente(proyecto_prime, "N/A", mes)

                                    if horas_opx:
                                        relation = float(horas_prime / horas_opx)
                                    else:
                                        relation = 0
                                    #print(project_default, mes, comp, personitas._nombre, "N/A", proyecto_prime, horas_opx, horas_prime, relation)
                                    
                                    #print(project_default, mes, comp, personitas._nombre, proyecto_opx, proyecto_prime, horas_opx, horas_prime, relation)
                                    line = (project_default, mes, comp, personitas._nombre, proyecto_opx, proyecto_prime, horas_opx, horas_prime, relation)
                                    unique_lines.add(line)  # Add the line to the set
                                    break


# Convertir el conjunto a una lista y ordenar
sorted_lines = sort_projects_by_month(unique_lines)
# Process and write the unique lines to Excel
for line in sorted_lines:
    #print('   '.join(map(str, line)))
    datos = [
        capitalizar_palabras(line[0]),  # project_default
        line[1],                        # mes
        convertir_mayusculas(line[2]),  # comp
        capitalizar_palabras(line[3]),  # personitas._nombre
        convertir_mayusculas(line[4]),  # proyecto_opx
        convertir_mayusculas(line[5]),  # proyecto_prime
        line[6],                        # horas_opx
        line[7],                        # horas_prime
        line[8]                         # relation
    ]

    for j in range(len(datos)):
        celda = hoja2.cell(row=num_fila_2, column=num_columna_2 + j, value=datos[j])
    num_fila_2 += 1  # Move to the next row for the next set of data
    cont += 1                      
            
                      


##############################################
#               Apply colors                #
##############################################
workbook.save(ruta_completa)
for i in range(cont):
    celda = hoja2.cell(row=i+TABLE_START_ROW+1, column = len(OUTPUT_HEADERS) + TABLE_START_COLUMN - 1)
    print(celda.value)
    if float(celda.value) > HIGH_PERCENTAGE:
        celda.fill = relleno_naranja
    elif float(celda.value) < LOW_PERCENTAGE:
        celda.fill = relleno_rojo
    else:
        celda.fill = relleno_verde


##############################################
#          Dar formato a la Hoja             #
##############################################

# Hacer una tabla en la hoja secundaria
existing_tables = hoja2.tables
table_name = 'Tabla_general'
 
# Verificar si la tabla ya existe y eliminarla si es necesario
if table_name in existing_tables:
    hoja2._tables.remove(existing_tables[table_name])


#print(((chr(TABLE_START_COLUMN+64)) + str(TABLE_START_ROW)) + ":" + ((chr( TABLE_START_COLUMN + len(OUTPUT_HEADERS) - 1 + 64 )) +str(num_fila_2 - 1)))
# Crear una nueva tabla
table = Table(
    displayName=table_name, 
    ref=(
        ((chr(TABLE_START_COLUMN+64)) + str(TABLE_START_ROW)) + 
        ":" + 
        ((chr( TABLE_START_COLUMN + len(OUTPUT_HEADERS) - 1 + 64 )) +str(num_fila_2 - 1))
        )
    )
style = TableStyleInfo(
    name=TABLE_STYLE, showFirstColumn=False,
    showLastColumn=False, showRowStripes=False, showColumnStripes=True)
table.tableStyleInfo = style
hoja2.add_table(table)

# Ajustar el ancho de las columnas en la hoja principal
for columna in hoja.iter_cols(min_col=2, max_col=3):
    longitud_maxima = 0
    columna_letra = columna[0].column_letter  # Obtener la letra de la columna
    for celda in columna:
        if celda.value:
            if celda.data_type == 'n':
                longitud_celda = 8
            else:
                # Calcular la longitud máxima del contenido en la columna
                longitud_celda = len(str(celda.value))
            
            # Actualizar la longitud máxima si es necesario
            if longitud_celda > longitud_maxima:
                longitud_maxima = longitud_celda
 
    # Establecer el ancho de la columna para ajustarlo al contenido más largo
    hoja.column_dimensions[columna_letra].width = longitud_maxima + 2

# Ajustar el ancho de las columnas en la hoja secundaria
for columna in hoja2.iter_cols(min_col=TABLE_START_COLUMN, max_col=TABLE_START_COLUMN+len(OUTPUT_HEADERS)-1):
    longitud_maxima = 0
    columna_letra = columna[0].column_letter  # Obtener la letra de la columna
    for celda in columna:
        if celda.value:
            if celda.data_type == 'n':
                longitud_celda = 8
            else:
                # Calcular la longitud máxima del contenido en la columna
                longitud_celda = len(str(celda.value))
            
            # Actualizar la longitud máxima si es necesario
            if longitud_celda > longitud_maxima:
                longitud_maxima = longitud_celda
 
    # Establecer el ancho de la columna para ajustarlo al contenido más largo
    hoja2.column_dimensions[columna_letra].width = longitud_maxima +2


workbook.save(ruta_completa)

if os.path.exists(ruta_completa):
    easygui.msgbox("Save completed, opening file...")
    os.startfile(ruta_completa)
else:
    easygui.msgbox(f'The file "{ruta_completa}" does not exist.')
