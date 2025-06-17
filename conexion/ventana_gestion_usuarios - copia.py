import sys,os, redshift_connector
import bcrypt
import mysql.connector
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QMessageBox, QTableWidget, QTableWidgetItem, QHBoxLayout, QHeaderView
)
from PySide6.QtGui import QFont, QIcon

class VentanaGestionUsuarios(QWidget):
      def __init__(self):
          super().__init__()
          self.personalizar_ventana()
          self.personalizar_componentes()
          #self.cargar_datos()

      def personalizar_ventana(self):
        self.setWindowTitle("Gestión de Usuarios")  # Título para la ventana
        self.setFixedSize(600, 400)  # Tamaño de la ventana ancho y altura
        #self.setStyleSheet("background-color: lightgray;")  # Color de fondo para la ventana

        # Cambiar el icono de la ventanapython9_ventana_menu_mihaita/cross1.png con una ruta absoluta que se crea a partir de una relativa
        ruta_relativa = "python14_ferreteria/cross1.png"
        ruta_absoluta = os.path.abspath(ruta_relativa)
        self.setWindowIcon(QIcon(ruta_absoluta))

      def personalizar_componentes(self):
          layout = QVBoxLayout()

          #Tabla de usuarios
          self.tabla = QTableWidget()
          self.tabla.setColumnCount(3)
          self.tabla.setHorizontalHeaderLabels(["CCONTRPS", "DCOMENTA"])
          self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

          # Campos para CRUD
          self.txt_nombre_usuario = QLineEdit()
          self.txt_nombre_usuario.setPlaceholderText("Nombre Usuario")

          self.txt_contrasena = QLineEdit()
          self.txt_contrasena.setPlaceholderText("Contraseña")

          self.txt_rol = QLineEdit()
          self.txt_rol.setPlaceholderText("Rol (Administrador/Cajero/Almacén)")

          # Botones para CRUD
          self.btn_agregar_usuario = QPushButton("Agregar Usuario")
          self.btn_agregar_usuario.clicked.connect(self.agregar_usuario)

          self.btn_editar_usuario = QPushButton("Editar Usuario")
          self.btn_editar_usuario.clicked.connect(self.editar_usuario)

          self.btn_eliminar_usuario = QPushButton("Eliminar Usuario")
          self.btn_eliminar_usuario.clicked.connect(self.eliminar_usuario)

          layout_botones = QHBoxLayout()
          layout_botones.addWidget(self.btn_agregar_usuario)
          layout_botones.addWidget(self.btn_editar_usuario)
          layout_botones.addWidget(self.btn_eliminar_usuario)

          layout.addWidget(self.tabla)
          layout.addWidget(self.txt_nombre_usuario)
          layout.addWidget(self.txt_contrasena)
          layout.addWidget(self.txt_rol)
          layout.addLayout(layout_botones)

          # Poner el administrador principal a la ventana(contenedor)
          self.setLayout(layout)

      def cargar_datos(self):
          conexion = self.obtener_conexion()
          if conexion != None:
             QMessageBox.information(self, "Ok", "Conexion")
             try:
                cursor = conexion.cursor()
                query = "SELECT ccontrps, dcomenta FROM ed_owner.t_ed_d_alarmasconfact limit 50" 
                cursor.execute(query)
                registros_lt = cursor.fetchall()
                self.tabla.setRowCount(len(registros_lt))
                for fila, registro_t in enumerate(registros_lt):
                    for columna, dato in enumerate(registro_t):
                        self.tabla.setItem(fila,columna, QTableWidgetItem(str(dato)))
             except Exception as e:
                 QMessageBox.critical(self, "Error", "Query Select")
          else:
             QMessageBox.critical(self, "Error", "Conexion")


      def obtener_conexion(self):
        conn = None
        try:
            conn: redshift_connector.Connection = redshift_connector.connect(
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
            conn = None
        return conn

      def agregar_usuario(self):
          conexion = self.obtener_conexion() 
          if conexion != None:
             nombre_usuario = self.txt_nombre_usuario.text()
             contrasena = self.txt_contrasena.text()
             rol = self.txt_rol.text()
             QMessageBox.information(self, "Ok", "Conexion")
             try:
                cursor = conexion.cursor()
                query = """INSERT INTO Usuario(nombre_usuario,contrasena,rol) 
                      VALUES (%s,%s,%s);"""
                cursor.execute(query,(nombre_usuario,self.encriptar_contrasena(contrasena),rol))
                conexion.commit()
                self.cargar_datos()
                QMessageBox.information(self, "Ok", "Query Insert")
             except Exception as e:
                QMessageBox.critical(self, "Error", "Query Insert")
          else:
             QMessageBox.critical(self, "Error", "Conexion")

      def encriptar_contrasena(self,contrasena):
          # CONVERTIR LA CONTRASEÑA A BYTES
          contrasena_byte = contrasena.encode()
          contrasena_hashed = bcrypt.hashpw(contrasena_byte, bcrypt.gensalt())
          return contrasena_hashed.decode()  
      
      def editar_usuario(self):
          conexion = self.obtener_conexion()
          if conexion != None:
             fila_seleccionada = self.tabla.currentRow()
             if fila_seleccionada != -1:
                nombre_usuario_buscar = self.tabla.item(fila_seleccionada,0).text()
               
                nombre_usuario_update = self.txt_nombre_usuario.text()
                contrasena_update = self.txt_contrasena.text()
                rol_update = self.txt_rol.text()

                if len(nombre_usuario_update) > 0 and len(contrasena_update) > 0 and \
                   len(rol_update) > 0:
                  
                   try:
                    cursor = conexion.cursor()
                    query = """UPDATE Usuario SET nombre_usuario = %s, contrasena = %s, rol = %s
                            WHERE nombre_usuario = %s;"""
                    cursor.execute(query,(nombre_usuario_update,self.encriptar_contrasena(contrasena_update),rol_update,nombre_usuario_buscar))
                    conexion.commit()
                    QMessageBox.information(self, "Ok", "Query Update")
                    self.cargar_datos()
                    self.limpiar_campos_usuario()
                   except Exception as e:
                       QMessageBox.critical(self, "Error", "Query Update")
                else:
                   QMessageBox.warning(self, "Warning", "Debe llenar todos los campos del usuario")    
             else:
                   QMessageBox.warning(self, "Warning", "Debe seleccionar usuario para editar") 
          else:
             QMessageBox.critical(self, "Error", "Conexion")

      def limpiar_campos_usuario(self):
          self.txt_nombre_usuario.clear()
          self.txt_contrasena.clear()
          self.txt_rol.clear()

      def eliminar_usuario(self):
          conexion = self.obtener_conexion()
          if conexion != None:
             fila_seleccionada = self.tabla.currentRow()
             if fila_seleccionada != -1:
                nombre_usuario_eliminar = self.tabla.item(fila_seleccionada,0).text()
                try:
                    cursor = conexion.cursor()
                    query = "DELETE FROM Usuario WHERE nombre_usuario = %s;"
                    cursor.execute(query,(nombre_usuario_eliminar,))
                    conexion.commit()
                    QMessageBox.information(self, "Ok", "Query Delete")
                    self.cargar_datos()
                    self.limpiar_campos_usuario()
                except Exception as e:
                    QMessageBox.critical(self, "Error", "Query Delete") 
             else:
                QMessageBox.warning(self, "Warning", "Debe seleccionar usuario para eliminar")  
          else:
            QMessageBox.critical(self, "Error", "Conexion")
             
# datetime.now()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ventana = VentanaGestionUsuarios()
    ventana.show()
    sys.exit(app.exec())
