import sys
import os
import redshift_connector
from PySide6.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
                               QMessageBox, QTableWidget, QTableWidgetItem, QHBoxLayout, QHeaderView, QFileDialog, QProgressBar)
from PySide6.QtGui import QFont, QIcon
import csv

class VentanaGestionUsuarios(QWidget):
    def __init__(self):
        super().__init__()
        self.personalizar_ventana()
        self.personalizar_componentes()
        self.cargar_datos()

    def personalizar_ventana(self):
        self.setWindowTitle("Extracción Contratos y sus Comentarios")
        self.setFixedSize(1200, 600)
        ruta_relativa = "python14_ferreteria/cross1.png"
        ruta_absoluta = os.path.abspath(ruta_relativa)
        self.setWindowIcon(QIcon(ruta_absoluta))

    def personalizar_componentes(self):
        layout = QVBoxLayout()

        # Tabla de usuarios
        self.tabla = QTableWidget()
        self.tabla.setColumnCount(16)
        self.tabla.setHorizontalHeaderLabels(["cemptitu", "finialar", "ffinalar", "cd_tp_alarma", "ccontrps", "id_crto_ext",
                                              "cd_cups_ext", "cd_sec_crto", "de_estado_crto", "lg_activo", "de_empr",
                                              "fh_alta_crto", "fh_baja_crto", "dcomenta", "de_usuario", "lg_origen"])
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        # Botones para CRUD
        self.btn_exportar_csv = QPushButton("Exportar a CSV")
        self.btn_exportar_csv.clicked.connect(self.exportar_a_csv)

        layout_botones = QHBoxLayout()
        layout_botones.addWidget(self.btn_exportar_csv)

        layout.addWidget(self.tabla)
        layout.addWidget(self.progress_bar)
        layout.addLayout(layout_botones)

        self.setLayout(layout)

    def exportar_a_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "Guardar CSV", "", "CSV Files (*.csv);;All Files (*)")
        if path:
            try:
                with open(path, mode='w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    # Escribir los encabezados
                    headers = [self.tabla.horizontalHeaderItem(i).text() for i in range(self.tabla.columnCount())]
                    writer.writerow(headers)
                    # Escribir los datos de la tabla
                    for row in range(self.tabla.rowCount()):
                        row_data = [self.tabla.item(row, col).text() if self.tabla.item(row, col) else '' for col in range(self.tabla.columnCount())]
                        writer.writerow(row_data)
                QMessageBox.information(self, "Éxito", "Datos exportados correctamente a CSV")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo exportar a CSV: {str(e)}")

    def cargar_datos(self):
        conexion = self.obtener_conexion()
        if conexion is not None:
            QMessageBox.information(self, "Ok", "Conexión exitosa")
            try:
                cursor = conexion.cursor()
                query = """SELECT a.cemptitu, a.finialar, a.ffinalar, a.cd_tp_alarma, a.ccontrps, b.id_crto_ext,
                                   b.cd_cups_ext, b.cd_sec_crto, b.de_estado_crto, b.lg_activo, b.de_empr,
                                   b.fh_alta_crto, b.fh_baja_crto, a.dcomenta, a.de_usuario, a.lg_origen
                            FROM ed_owner.t_ed_d_alarmasconfact a
                            INNER JOIN ed_owner.t_ed_f_uvcrtos b ON a.ccontrps = b.id_crto_ext
                            WHERE a.lg_origen = 'SAP'
                            AND b.lg_activo = 'S'
                            AND a.ffinalar <> '99991231'
                            AND b.cd_cups_ext IS NOT NULL"""
                cursor.execute(query)
                registros_lt = cursor.fetchall()
                self.tabla.setRowCount(len(registros_lt))
                
                # Actualizar la barra de progreso
                total_registros = len(registros_lt)
                self.progress_bar.setMaximum(total_registros)
                
                for fila, registro_t in enumerate(registros_lt):
                    for columna, dato in enumerate(registro_t):
                        self.tabla.setItem(fila, columna, QTableWidgetItem(str(dato)))
                    self.progress_bar.setValue(fila + 1)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error en la consulta: {str(e)}")
            finally:
                conexion.close()
        else:
            QMessageBox.critical(self, "Error", "Error en la conexión")

    def obtener_conexion(self):
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
            QMessageBox.critical(self, "Error", f"Error en la conexión: {str(e)}")
        return conn

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ventana = VentanaGestionUsuarios()
    ventana.show()
    sys.exit(app.exec())
           