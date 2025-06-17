
import redshift_connector
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

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
        messagebox.showerror("Error de conexión", f"Error al conectar: {e}")
        conn = None
    return conn

def leer_cd_fiscal_fact(file_path):
    cd_fiscal_fact_list = []
    with open(file_path, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            cd_fiscal_fact_list.append(row[0])
    return cd_fiscal_fact_list

def buscar_en_tabla(conn, cd_fiscal_fact_list, progress_var):
    query = f"""
    SELECT cd_fiscal_fact, de_est_oblg, fh_emision_fact, fh_puesta_cobro, fh_cobro,
           lg_oblg_cobrada, fh_vencimiento, fh_cobro_real, fh_ini_vencimiento
    FROM ed_owner.t_ed_h_obligaciones
    WHERE cd_fiscal_fact IN ({','.join("'" + str(x) + "'" for x in cd_fiscal_fact_list)});
    """
    cursor = conn.cursor()
    cursor.execute(query)
    
    results = []
    total_rows = cursor.rowcount
    processed_rows = 0
    
    for row in cursor:
        results.append(row)
        processed_rows += 1
        progress_var.set((processed_rows / total_rows) * 100)
    
    return results

def mostrar_resultados(results):
    root = tk.Tk()
    root.title("Resultados de la Consulta")

    tree = ttk.Treeview(root, columns=('cd_fiscal_fact', 'de_est_oblg', 'fh_emision_fact', 'fh_puesta_cobro', 'fh_cobro',
                                       'lg_oblg_cobrada', 'fh_vencimiento', 'fh_cobro_real', 'fh_ini_vencimiento'), show='headings')
    tree.heading('cd_fiscal_fact', text='cd_fiscal_fact')
    tree.heading('de_est_oblg', text='de_est_oblg')
    tree.heading('fh_emision_fact', text='fh_emision_fact')
    tree.heading('fh_puesta_cobro', text='fh_puesta_cobro')
    tree.heading('fh_cobro', text='fh_cobro')
    tree.heading('lg_oblg_cobrada', text='lg_oblg_cobrada')
    tree.heading('fh_vencimiento', text='fh_vencimiento')
    tree.heading('fh_cobro_real', text='fh_cobro_real')
    tree.heading('fh_ini_vencimiento', text='fh_ini_vencimiento')

    for row in results:
        tree.insert('', tk.END, values=row)

    tree.pack(expand=True, fill='both')
    root.mainloop()

def exportar_csv(self):
    path = "C:\\temp\\datos.csv"
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
        QMessageBox.information(self, "Éxito", "Datos exportados correctamente a CSV en C:\\temp")
    except Exception as e:
        QMessageBox.critical(self, "Error", f"No se pudo exportar a CSV: {str(e)}")



def seleccionar_fichero(progress_var):
    file_path = filedialog.askopenfilename(title="Seleccionar fichero de entrada", filetypes=[("CSV files", "*.csv")])
    if file_path:
        # Leer los valores de cd_fiscal_fact del fichero de entrada
        cd_fiscal_fact_list = leer_cd_fiscal_fact(file_path)

        # Obtener la conexión a la base de datos
        conn = obtener_conexion()

        if conn:
            # Buscar en la tabla y obtener los resultados
            results = buscar_en_tabla(conn, cd_fiscal_fact_list, progress_var)
            
            # Mostrar los resultados en una ventana
            mostrar_resultados(results)
            
            # Cerrar la conexión
            conn.close()

# Crear la ventana principal
root = tk.Tk()
root.title("Programa de Consulta")

# Variable de progreso
progress_var = tk.DoubleVar()

# Barra de progreso
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.pack(pady=20)

# Botón para seleccionar el fichero de entrada
btn_seleccionar_fichero = tk.Button(root, text="Seleccionar Fichero de Entrada", command=lambda: seleccionar_fichero(progress_var))
btn_seleccionar_fichero.pack(pady=20)

try:
    root.mainloop()
except Exception as e:
    messagebox.showerror("Error de aplicación", f"Error en la aplicación: {e}")