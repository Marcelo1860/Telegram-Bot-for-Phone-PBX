import telepot
import pandas as pd
import os
import time

# cargar los datos del archivo Excel en un DataFrame
df = pd.read_excel('ficha_final_cliente.xlsx', sheet_name='hoja1')


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
    
    # leer los datos del DataFrame
    
    # enviar los datos al usuario
    bot.sendMessage(chat_id, 'se registraron los datos en el excel!')
    
def agregar_datos(df, nuevos_datos):
    # crear un nuevo DataFrame con los nuevos datos
    nuevos_df = pd.DataFrame(nuevos_datos, columns=df.columns)
    
    # concatenar el nuevo DataFrame con el original
    df = pd.concat([df, nuevos_df], ignore_index=True)
    
    # devolver el DataFrame actualizado
    return df

def guardar_datos(df, archivo):
    # guardar el DataFrame en un archivo Excel
    writer = pd.ExcelWriter(archivo, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='hoja1', index=False)
    writer.save()
    

# obtener el token del bot de Telegram
token = '6188339712:AAF04J6tDRs3lOAtDVYb7EUig6Iz-CiK_40'

# crear una instancia del bot
bot = telepot.Bot(token)


# definir una función para manejar los mensajes recibidos por el bot
def handle_message(msg):
    # obtener el chat_id del usuario
    chat_id = msg['chat']['id']
    
    # verificar si el usuario envió el comando para agregar nuevos datos
    if msg['text'].startswith('/agregar'):
        # obtener los nuevos datos como una lista de listas
        nuevos_datos = [msg['text'][9:].split()]
        
        # enviar los nuevos datos al usuario
        enviar_datos(chat_id, nuevos_datos, df)

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


