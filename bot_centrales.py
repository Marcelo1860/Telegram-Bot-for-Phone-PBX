import telepot
import pandas as pd
import os
import time
from PIL import Image, ImageDraw, ImageFont
import textwrap

# cargar los datos del archivo Excel en un DataFrame
df = pd.read_excel('ficha_final_cliente.xlsx', sheet_name='hoja1')
first = True
contador = -1
nombres_columnas = df.columns
lista_nombres_columnas = list(df.columns)

def leer_datos():
    # obtener los datos del DataFrame
    datos = df.to_string(index=False)
    
    # devolver los datos como una cadena de texto
    return datos

def enviar_datos(chat_id, nuevos_datos, df):
    # agregar los nuevos datos al DataFrame
    
    print('DataFrame antes de agregar nuevos datos:')
    print(df)
    
    df = agregar_datos(df, nuevos_datos)
    
    print('DataFrame después de agregar nuevos datos:')
    print(df)
    
    # guardar el DataFrame en el archivo Excel
    guardar_datos(df, 'ficha_final_cliente.xlsx')
    
    print('DataFrame después de guardar nuevos datos:')
    print(df)
        
def agregar_datos(df, nuevos_datos):
    # crear un nuevo DataFrame con los nuevos datos
    #nuevos_df = pd.DataFrame(nuevos_datos, columns=df.columns)
    
    # concatenar el nuevo DataFrame con el original
    df = pd.concat([df, nuevos_datos], ignore_index=True)
    
    # devolver el DataFrame actualizado
    return df

def guardar_datos(df, archivo):
    # guardar el DataFrame en un archivo Excel
    writer = pd.ExcelWriter(archivo, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='hoja1', index=False)
    writer.save()

# Definir la función que crea la imagen con los datos
def crear_imagen(df):
    # Crear una imagen en blanco
    imagen = Image.new('RGB', (800, 1600), color='white')
    
    # Crear un objeto Draw para dibujar en la imagen
    draw = ImageDraw.Draw(imagen)
    
    font_column = ImageFont.truetype('arial.ttf', 16)
    font_value = ImageFont.truetype('arial.ttf', 14)

    # Dibujar los textos en la imagen
    x_column, x_value = 50, 200
    y_start = 50
    y_step = 30
    prev_lines = 0
    for i, column in enumerate(df.columns):
        # Dibujar el nombre de la columna
        draw.text((x_column, y_start + y_step * 2 * i + prev_lines*y_step), column, fill='blue', font=font_column)
        
        # Obtener el valor de la última fila de la columna
        value = str(df.iloc[-1][column])
        
        # Dividir el valor en varias líneas si es necesario
        lines = textwrap.wrap(value, width=90)
        num_lines = len(lines)
        
        # Dibujar cada línea de texto
        for j, line in enumerate(lines):
            draw.text((x_column, y_start + y_step * (2*i+1) +(prev_lines+j)*y_step), line, fill='black', font=font_value)

        # Actualizar el número de líneas previas
        prev_lines = prev_lines +  num_lines -1    

    # Guardar la imagen en un archivo temporal
    imagen_path = 'temp.pdf'
    imagen.save(imagen_path)
    
    # Devolver el path de la imagen
    return imagen_path
    
    

# obtener el token del bot de Telegram
token = '6188339712:AAF04J6tDRs3lOAtDVYb7EUig6Iz-CiK_40'

# crear una instancia del bot
bot = telepot.Bot(token)


# definir una función para manejar los mensajes recibidos por el bot
def handle_message(msg):
    # obtener el chat_id del usuario

    chat_id = msg['chat']['id']
    
    # verificar si el usuario envió el comando para agregar nuevos datos
    command = msg['text']

    global contador

    print (df)

    if command == '/agregar':
        bot.sendMessage(chat_id, 'agregue la fecha del trabajo en formato DD/MM/AA')
    elif command == '/comprobante_foto':
        #fila = df.iloc[-1].tolist()
        imagen_path = crear_imagen(df)
        with open(imagen_path, 'rb') as imagen:
            bot.sendPhoto(chat_id, imagen)
        contador = -2
    elif command == '/comprobante_pdf':
        #fila = df.iloc[-1].tolist()
        imagen_path = crear_imagen(df)
        # enviar el archivo pdf al usuario
        with open('temp.pdf', 'rb') as f:
            bot.sendDocument(chat_id, f)
        contador = -2
    elif command == '/excel':
        with open('ficha_final_cliente.xlsx', 'rb') as f:
            bot.sendDocument(chat_id, document=f)
        contador = -2
    elif command == '/limpiar':
        df.drop(df.index, inplace=True)
        guardar_datos(df, 'ficha_final_cliente.xlsx')
        contador = -2
    elif contador == -1:
        bot.sendMessage(chat_id, 'necesita poner /agregar antes de poder hacer algo')
        return
    elif contador == 0: 
        new_row = pd.DataFrame({'Fecha': [command], 'Nombre cliente': [0], 'Modelo PBX': [0],'S/N': [command], 'Falla acusada': [0], 'Diagnostico': [0],'Resolucion': [0], 'PV': [0] })
        enviar_datos(chat_id,new_row,df)
        bot.sendMessage(chat_id, 'agregue {}'.format(lista_nombres_columnas[contador+1]))
    elif contador > 0:
        df.iloc[-1,contador]=command
        guardar_datos(df, 'ficha_final_cliente.xlsx')
        if contador < 7:
            bot.sendMessage(chat_id, 'agregue {}'.format(lista_nombres_columnas[contador+1]))

    contador +=1
    if contador == 8:
        bot.sendMessage(chat_id, 'todos los datos han sido agregados')
        bot.sendMessage(chat_id, 'para obtener comprobante mande /comprobante_foto o /comprobante_pdf')
        contador = -1



# agregar un manejador de eventos para el bot
bot.message_loop(handle_message)

# esperar a recibir mensajes
print('Esperando mensajes...')
while True:
    time.sleep(10)
    df = pd.read_excel('ficha_final_cliente.xlsx', sheet_name='hoja1')
    pass

# guardar el DataFrame en el archivo Excel al finalizar el programa
guardar_datos(df, 'ficha_final_cliente.xlsx')


