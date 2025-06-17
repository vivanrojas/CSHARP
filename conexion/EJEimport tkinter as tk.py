import tkinter as tk
from tkinter import Menu
import redshift_connector

# Definir la función obtener_conexion
def obtener_conexion():
    conn = None
    try:
        conn = redshift_connector.connect(
            iam=True,
            database='dtmco_pro',
            host='loib-ap02260-prod-rshift-master.cwevuvdkfkey.eu-central-1.redshift.amazonaws.com',
            credentials_provider='BrowserAzureOAuth2CredentialsProvider',
            scope='api://enel.com/b007de93-d931-45d4-8b57-819f4ea1ce4d/jdbc_login',
            cluster_identifier='loib-ap02260-prod-rshift-master',
            region='eu-central-1',
            listen_port=7890,
            idp_response_timeout=50,
            idp_tenant='d539d4bf-5610-471a-afc2-1c76685cfefa',
            client_id='b007de93-d931-45d4-8b57-819f4ea1ce4d',
        )
    except Exception as e:
        print(f"Error al conectar: {e}")
        conn = None
    return conn

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Ventana con Submenú")
ventana.geometry("400x300")

# Crear la barra de menú
barra_menus = Menu(ventana)
ventana.config(menu=barra_menus)

# Crear el menú "Opciones"
menu_archivo = Menu(barra_menus, tearoff=0)
barra_menus.add_cascade(label="Opciones", menu=menu_archivo)

# Añadir opciones al menú "Opciones"
menu_archivo.add_command(label="Conexion", command=obtener_conexion)
menu_archivo.add_command(label="Facturacion")
menu_archivo.add_separator()
menu_archivo.add_command(label="Salir", command=ventana.quit)

# Crear un submenú dentro del menú "Opciones"
sub_menu = Menu(menu_archivo, tearoff=0)
menu_archivo.add_cascade(label="Submenú", menu=sub_menu)

# Añadir opciones al submenú
sub_menu.add_command(label="Opción 1")
sub_menu.add_command(label="Opción 2")

# Crear el menú "Editar"
menu_editar = Menu(barra_menus, tearoff=0)
barra_menus.add_cascade(label="Editar", menu=menu_editar)
menu_editar.add_command(label="Cortar")
menu_editar.add_command(label="Copiar")
menu_editar.add_command(label="Pegar")

# Crear el menú "Ayuda"
menu_ayuda = Menu(barra_menus, tearoff=0)
barra_menus.add_cascade(label="Ayuda", menu=menu_ayuda)
menu_ayuda.add_command(label="Acerca de")
menu_ayuda.add_command(label="Ayuda")

# Ejecutar la ventana
ventana.mainloop()