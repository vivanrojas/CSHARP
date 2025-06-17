
import redshift_connector

def obtener_conexion():
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

def ejemplo1():
   conn = obtener_conexion()
   if conn != None:
        print("OK: CONEXION")
        cursor = conn.cursor()
        try:
            query = "SELECT * FROM ed_owner.t_ed_contratos_ml limit 50"
            cursor.execute(query)
            resultados_lt = cursor.fetchall()
            for resultado_t in resultados_lt:
                cextctr = resultado_t
                print(cextctr[0])
            print("OK: QUERY")
        except Exception as e:
            print("ERROR: QUERY")
   else:
        print("ERROR: CONEXION")

def ejemplo2():
   conn = obtener_conexion()
   if conn != None:
        print("OK: CONEXION")
        cursor = conn.cursor()
        try:
            query = "CREATE TABLE operaciones_se_medida_owner.Probando (id int, nombre varchar(30))"
            cursor.execute(query)
            print("OK: QUERY")
        except Exception as e:
            print("ERROR: QUERY ",e)
   else:
        print("ERROR: CONEXION")

ejemplo1()


'''
cursor: redshift_connector.Cursor = conn.cursor()
cursor.execute("select * from stv_wlm_query_state")
result: tuple = cursor.fetchall()
print(result)
'''
