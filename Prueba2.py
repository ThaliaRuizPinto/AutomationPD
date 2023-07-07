import openpyxl
import openpyxl.utils
import time

start_time = time.time()

# Configuración openpy:
archivo_excel = "C:\\Users\\thali\\Downloads\\MARCA PD 010723.xlsx"
libro = openpyxl.load_workbook(archivo_excel)
print(libro.sheetnames)

# LIBROS DE DATOS
ws_bd_activos = libro["BD ACTIVOS"]
ws_product_owner = libro["PRODUCT OWNER"]
ws_segmento_digital = libro["SEGMENTO DIGITAL "]
ws_perfil_digital = libro["PERFIL DIGITAL"]
ws_reglas_iniciales = libro["REGLAS INICIALES"]

# Valores BD ACTIVOS
area = ws_bd_activos["L2:L23292"]
servicio = ws_bd_activos["N2:N23292"]
uo = ws_bd_activos["P2:P23292"]
llave = ws_bd_activos["CW2:CW23292"]
division = ws_bd_activos["J2:J23292"]
Grupo = ws_bd_activos["AA2:AA23292"]
Funcion = ws_bd_activos["AC2:AC23292"]
Subgrupo = ws_bd_activos["AB2:AB23292"]

# Celdas de escritura
campo_po = ws_bd_activos["CZ2:CZ23292"]
campo_segmento = ws_bd_activos["CX2:CX23292"]
campo_perfil = ws_bd_activos["DA2:DA23292"]
campo_RI = ws_bd_activos["CY2:CY23292"]

# Valores SEGMENTO ATENCION DIGITAL
segmento_digital_division = ws_segmento_digital["A2:A4"]
segmento_digital_area = ws_segmento_digital["C2:C29"]
segmento_digital_uo = ws_segmento_digital["G2:G46"]

# Valores REGLAS INICIALES
Llaves_proveedores = ws_reglas_iniciales["A2:A16"]
Llaves_ADHOC = ws_reglas_iniciales["A43:A75"]

#Valores EXCEPCIONES
codigo_funcion_practi = ws_reglas_iniciales["D25"]
grupos_X = ws_reglas_iniciales["C28:C29"]
subgrupos_X = ws_reglas_iniciales["D34:D38"]
division_X = ws_reglas_iniciales["B20"]

# Valores PRODUCT OWNER
area_po = ws_product_owner["F2:F10"]
servicio_po = ws_product_owner["H11:H17"]
uo_po = ws_product_owner["J18:J36"]
uo_po_no_si = ws_product_owner["J37:J150"]
uo_po_no_no = ws_product_owner["J151:J156"]


# Valores PERFIL DIGITAL
area_pF = ws_perfil_digital["A2:A7"]
servicio_pF = ws_perfil_digital["G8:G18"]
uo_pF_si = ws_perfil_digital["A19:A706"]
uo_pF_no = ws_perfil_digital["I707:I757"]
llave_si = ws_perfil_digital["A758:A816"]
llave_no = ws_perfil_digital["A817:A820"]


# LOGICA SEGMENTO ATENCIÓN DIGITAL --------------------------------------------------------------------
for i in range(len(campo_segmento)):
    if division[i][0].value in [cell[0].value for cell in segmento_digital_division]:
        campo_segmento[i][0].value = "SI"
    elif area[i][0].value in [cell[0].value for cell in segmento_digital_area]:
        campo_segmento[i][0].value = "SI"
    elif llave[i][0].value in [cell[0].value for cell in segmento_digital_uo]:
        campo_segmento[i][0].value = "SI"
    else:
        campo_segmento[i][0].value = "NO"

print("Segmento Atención Digital -------------- COMPLETE")



# EXCEPCIONES Y REGLAS INICIALES
for i in range(len(campo_RI)):
    if Funcion[i][0].value == codigo_funcion_practi.value and campo_segmento[i][0].value == "SI":
        campo_RI[i][0].value = "NO"
    elif Grupo[i][0].value in [cell[0].value for cell in grupos_X] and campo_segmento[i][0].value == "SI":
        campo_RI[i][0].value = "NO"
    elif division[i][0].value == division_X.value and campo_segmento[i][0].value == "SI":
        campo_RI[i][0].value = "NO"
    elif Subgrupo[i][0].value in [cell[0].value for cell in subgrupos_X] and campo_segmento[i][0].value == "SI":
        campo_RI[i][0].value = "SI"
    else:
        if llave[i][0].value in [cell[0].value for cell in Llaves_proveedores]:
            campo_RI[i][0].value = "SI"
        elif llave[i][0].value in [cell[0].value for cell in Llaves_ADHOC]:
            campo_RI[i][0].value = "SI"
        else:
            campo_RI[i][0].value = "NO"



# LOGICA PRODUCT OWNER --------------------------------------------------------------------------------
for i in range(len(Grupo)):
    if Grupo[i][0].value == "Tribu PO":
        if campo_segmento[i][0].value == "SI":
            if area[i][0].value in [cell[0].value for cell in area_po]:
                campo_po[i][0].value = "SI"
                
            elif servicio[i][0].value in [cell[0].value for cell in servicio_po]:
                campo_po[i][0].value = "SI"
               
            elif uo[i][0].value in [cell[0].value for cell in uo_po]:
                campo_po[i][0].value = "SI"
            else: 
                campo_po[i][0].value = "NO"
        elif campo_segmento[i][0].value == "NO":
            if uo[i][0].value in [cell[0].value for cell in uo_po_no_si]:
                campo_po[i][0].value = "SI"
            elif uo[i][0].value in [cell[0].value for cell in uo_po_no_no]:
                campo_po[i][0].value = "NO"
            else:
                campo_po[i][0].value = "NO"
        else:
            campo_po[i][0].value = "NO"
    else:
        campo_po[i][0].value = "NO"

print("Product Owner -------------- COMPLETE")

# LOGICA PERFIL DIGITAL ---------------------------------------------------------------------------------
for i in range(len(campo_perfil)):
    if area[i][0].value in [str(cell[0].value) for cell in area_pF]:
        campo_perfil[i][0].value = "SI"
    elif servicio[i][0].value in [cell[0].value for cell in servicio_pF]:
        campo_perfil[i][0].value = "SI"
    elif uo[i][0].value in [cell[0].value for cell in uo_pF_si]:
        campo_perfil[i][0].value = "SI"
    elif uo[i][0].value in [cell[0].value for cell in uo_pF_no]:
        campo_perfil[i][0].value = "NO"
    elif llave[i][0].value in [cell[0].value for cell in llave_si]:
        campo_perfil[i][0].value = "SI"
    elif llave[i][0].value in [cell[0].value for cell in llave_no]:
        campo_perfil[i][0].value = "NO"
    else:
        campo_perfil[i][0].value = "NO"
        

# Segunda validación de llave
for i in range(len(campo_perfil)):
    if llave[i][0].value in [cell[0].value for cell in llave_si]:
        campo_perfil[i][0].value = "SI"
    elif llave[i][0].value in [cell[0].value for cell in llave_no]:
        campo_perfil[i][0].value = "NO"

print("Perfil Digital -------------- COMPLETE")


libro.save(archivo_excel)
print("GUARDADO")
end_time = time.time()
execution_time = end_time - start_time
print("Tiempo de ejecución:", execution_time, "segundos")
