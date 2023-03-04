import telepot
import pandas as pd
import os
import time
from PIL import Image, ImageDraw, ImageFont

# cargar los datos del archivo Excel en un DataFrame
df = pd.read_excel('ficha_final_cliente.xlsx', sheet_name='hoja1')
first = True
contador = -1
nombres_columnas = df.columns
lista_nombres_columnas = list(df.columns)
# Definir la fuente para el texto
font = ImageFont.truetype('arial.ttf', size=10)

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
def crear_imagen(fila):
    # Crear una imagen en blanco
    imagen = Image.new('RGB', (400, 400), color='white')
    
    # Crear un objeto Draw para dibujar en la imagen
    draw = ImageDraw.Draw(imagen)
    
    # Escribir los datos en la imagen
    mensaje = '\n'.join([str(dato) for dato in fila])
    draw.text((10, 10), mensaje, font=font, fill='black')
    
    # Guardar la imagen en un archivo temporal
    imagen_path = 'temp.jpg'
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
    elif command == '/comprobante':
        fila = df.iloc[-1].tolist()
        imagen_path = crear_imagen(fila)
        with open(imagen_path, 'rb') as imagen:
            bot.sendPhoto(chat_id, imagen)
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
        bot.sendMessage(chat_id, 'para obtener comprobante mande /comprobante')
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


