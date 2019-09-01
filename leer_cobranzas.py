import pandas as pd
import numpy as np
import xlsxwriter
from datetime import datetime



#df = pd.read_excel (r'COBRANZAS.xls') #for an earlier version of Excel, you may need to use the file extension of 'xls'

#print (df)
print ('*** Cuentas Contables *****')
print ('Tipo           DEBE                               HABER')
print ('-----------------------------------------------------------')
print ('Ingreso         1.01.001.002.002          1.01.003.001.999')
print ('Egreso          6.02.020.001.999          1.01.001.002.002')
print ('Comisiones      6.02.010.001.001          1.01.001.002.002')

mensaje_ingreso='REGISTRO COBRANZAS DEL MES '
mensaje_egreso='REGISTRO COBRANZAS DEL MES '
mensaje_comision='REGISTRO COBRANZAS DEL MES '
mensaje2='REGISTRO COBRANZAS FACTURA '

#print(df.sheet_names)
egresos = []
ingresos = []
comisiones = []
mes = str(input("Indicar mes a procesar:    "))  
mes =mes.upper()
comprobante = str(input("Digame el Numero de Comprobante:    "))  

ingreso_cuenta_deudora = str(input("Indicar Ingreso cuenta_deudora:    ")) 
if ingreso_cuenta_deudora =='':
    ingreso_cuenta_deudora='1.01.001.002.002'

ingreso_cuenta_acreedora = str(input("Indicar Ingreso cuenta_acreedora:    "))  
if ingreso_cuenta_acreedora =='':
    ingreso_cuenta_acreedora='1.01.003.001.999'



egreso_cuenta_deudora = str(input("Indicar egreso cuenta_deudora:    ")) 
if egreso_cuenta_deudora =='':
    egreso_cuenta_deudora='6.02.020.001.999'

egreso_cuenta_acreedora = str(input("Indicar egreso cuenta_acreedora:    "))  
if egreso_cuenta_acreedora =='':
    egreso_cuenta_acreedora='1.01.001.002.002'


comision_cuenta_deudora = str(input("Indicar comision cuenta_deudora:    ")) 
if comision_cuenta_deudora =='':
    comision_cuenta_deudora='6.02.010.001.001'

comision_cuenta_acreedora = str(input("Indicar comision cuenta_acreedora:    "))  
if comision_cuenta_acreedora =='':
    comision_cuenta_acreedora='1.01.001.002.002'

xlsx = pd.ExcelFile('COBRANZAS.xls')
df = pd.read_excel(xlsx,  mes)

for index, row in df.iterrows():
    linead = []
    lineah = []
    fecha = row['Fecha']
    fecha=fecha.strftime("%d/%m/%Y")
    monto = row['Monto']
    descripcion = row['Descripción']
    balance = row['Balance']
    referencia = row['Referencia']
    #Linea D
    linead.append(comprobante)
    linead.append(fecha)
    
    if monto<0 and descripcion!='COMISION TRF OTROS BCOS':
        linead.append(mensaje_egreso)
        linead.append(egreso_cuenta_deudora)
    elif monto>0 and descripcion!='COMISION TRF OTROS BCOS' :
        linead.append(mensaje_ingreso)
        linead.append(ingreso_cuenta_deudora)
    elif descripcion=='COMISION TRF OTROS BCOS':
        linead.append(mensaje_comision)
        linead.append(comision_cuenta_deudora)

    linead.append(descripcion)
    linead.append('D')
    linead.append(monto)
    linead.append(fecha)
    linead.append(referencia)
    linead.append(descripcion)
    linead.append(3)
    #Linea H
    lineah.append(comprobante)
    lineah.append(fecha)
   

    
    if monto<0 and descripcion!='COMISION TRF OTROS BCOS':
        lineah.append(mensaje_egreso)
        lineah.append(egreso_cuenta_acreedora)
    elif monto>0 and descripcion!='COMISION TRF OTROS BCOS' :
        lineah.append(mensaje_ingreso)
        lineah.append(ingreso_cuenta_acreedora)
    elif descripcion=='COMISION TRF OTROS BCOS':
        lineah.append(mensaje_comision)
        lineah.append(comision_cuenta_acreedora)

    lineah.append(descripcion)
    lineah.append('H')
    lineah.append(monto)
    lineah.append(fecha)
    lineah.append(referencia)
    lineah.append(descripcion)
    lineah.append(3)
    
    if monto<0 and descripcion!='COMISION TRF OTROS BCOS':
        egresos.append(linead)
        egresos.append(lineah)
    elif monto>0 and descripcion!='COMISION TRF OTROS BCOS' :
        ingresos.append(linead)
        ingresos.append(lineah)
    elif descripcion=='COMISION TRF OTROS BCOS':
        comisiones.append(linead)
        comisiones.append(lineah)


# Create a Pandas dataframe from some data.
df_egresos = pd.DataFrame(egresos)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer_egresos = pd.ExcelWriter(mes+'_egresos.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df_egresos.to_excel(writer_egresos, sheet_name=mes+'_Egresos')

# Close the Pandas Excel writer and output the Excel file.
writer_egresos.save()

df_ingresos = pd.DataFrame(ingresos)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer_ingresos = pd.ExcelWriter(mes+'_ingresos.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df_ingresos.to_excel(writer_ingresos, sheet_name=mes+'_Ingresos')

# Close the Pandas Excel writer and output the Excel file.
writer_ingresos.save()

df_comsiones = pd.DataFrame(comisiones)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer_comsiones = pd.ExcelWriter(mes+'_comsiones.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df_comsiones.to_excel(writer_comsiones, sheet_name=mes+'_Comisiones')

# Close the Pandas Excel writer and output the Excel file.
writer_comsiones.save()