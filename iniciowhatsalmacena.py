# -*- coding: utf-8 -*-
"""
Created on Thu Dec 21 14:35:19 2023

@author: Bryan Hernandez
"""

import pandas as pd
import pywhatkit as pwk
from datetime import datetime, timedelta

data = pd.read_excel("Clientes.xlsx", sheet_name='Ventas')
data.head(3)

# Abre un archivo en modo de escritura (se crea el archivo si no existe)

for i in range(len(data)):
        celular = data.loc[i, 'Celular'].astype(str)
        nombre = data.loc[i, 'Nombre']
        producto = data.loc[i, 'Producto']
        imagen = data.loc[i, 'imagen']
        saludo = data.loc[i, 'saludo']

        # Obtener la hora actual en cada iteración
        hora_actual = datetime.now()

        # Sumar 45 segundos a la hora actual en cada iteración
        hora_programada = hora_actual + timedelta(seconds=35)

        # Crear mensaje personalizado
        mensaje = f"{saludo} {nombre} \n {producto}"

        # Preparacion de envío de mensaje
        pwk.sendwhats_image("+" + celular, "Images/" + imagen, mensaje, 15, True, 2)

        # Imprime el mensaje en la consola y guarda en el archivo
        message_info = f"Message Sent to {nombre} at {hora_programada.time()}"

