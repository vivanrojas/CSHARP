using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.servidores
{
    class RedShiftServer
    {
        private Dictionary<string, string> lista_parametros;

        public OdbcConnection con;

        utilidades.Param p;

        public enum Entornos
        {
            PROD,
            UAT4,
            PROD_CONTRATACION,
            PROD_FACTURACION,
            PROD_MEDIDA
        }

        public RedShiftServer(Entornos entorno)
        {
            try
            {
                p = new utilidades.Param("parameters", MySQLDB.Esquemas.AUX);
                con = new OdbcConnection(GetConnectionString(entorno));
                //con.ConnectionTimeout = 1000;
                con.Open();
            }
            catch (OdbcException e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private string GetConnectionString(Entornos entorno)
        {
            string server;
            string port;
            string ddbb;
            string user;
            string pass;

            try
            {
                
                // Cargamos valores por defecto, iguales para todos los esquemas del entorno de producción, excepto el usuario y contraseña que actualizamos dependiendo del entorno seleccionado
                server = p.GetValue("redshift_server", DateTime.Now, DateTime.Now); // A 23-04-2025: "loib-ap02260-prod-rshift-master.cwevuvdkfkey.eu-central-1.redshift.amazonaws.com";
                ddbb = p.GetValue("redshift_ddbb", DateTime.Now, DateTime.Now);
                port = p.GetValue("redshift_port", DateTime.Now, DateTime.Now);
                
                //user = p.GetValue("redshift_user", DateTime.Now, DateTime.Now);
                //Encriptamos la contraseña en base de datos y modificamos para desencriptar aquí
                //string pass = p.GetValue("redshift_pass", DateTime.Now, DateTime.Now);
                //pass = utilidades.FuncionesTexto.Decrypt(p.GetValue("redshift_pass", DateTime.Now, DateTime.Now), true);

                switch (entorno)
                {
                    case Entornos.PROD:
                        user = p.GetValue("redshift_user", DateTime.Now, DateTime.Now);
                        pass = utilidades.FuncionesTexto.Decrypt(p.GetValue("redshift_pass", DateTime.Now, DateTime.Now), true);
                        break;
                    case Entornos.PROD_CONTRATACION:
                        user = p.GetValue("redshift_user_contratacion", DateTime.Now, DateTime.Now);
                        pass = utilidades.FuncionesTexto.Decrypt(p.GetValue("redshift_pass_contratacion", DateTime.Now, DateTime.Now), true);
                        break;
                    case Entornos.PROD_FACTURACION:
                        user = p.GetValue("redshift_user_facturacion", DateTime.Now, DateTime.Now);
                        pass = utilidades.FuncionesTexto.Decrypt(p.GetValue("redshift_pass_facturacion", DateTime.Now, DateTime.Now), true);
                        break;
                    case Entornos.PROD_MEDIDA:
                        user = p.GetValue("redshift_user_medida", DateTime.Now, DateTime.Now);
                        pass = utilidades.FuncionesTexto.Decrypt(p.GetValue("redshift_pass_medida", DateTime.Now, DateTime.Now), true);
                        break;
                    case Entornos.UAT4:
                        server = "loib-ap02260-test-rshift-master-2.cluashiidwwf.eu-central-1.redshift.amazonaws.com";
                        ddbb = " dtmco_pru2";
                        user = "pruebas_dtmco_pru";
                        pass = "Pru3b4s_pre";
                        break;
                    default:
                        user = p.GetValue("redshift_user", DateTime.Now, DateTime.Now);
                        pass = utilidades.FuncionesTexto.Decrypt(p.GetValue("redshift_pass", DateTime.Now, DateTime.Now), true);
                        break;

                }

               // Console.WriteLine(String.Format("Datos de conexión a RedShift, entorno {4}: \r\nServidor = {0} \r\nPuerto = {1} \r\nBase de datos = {2} \r\nUsuario = {3}", server, port, ddbb, user, entorno));

                return "Driver={Amazon Redshift (x64)};" +
                        String.Format("Server={0};Database={1};" +
                        "UID={2};PWD={3};Port={4};KeepAlive={5}",
                        server, ddbb, user, pass, port, 1);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }


        }

        public void CloseConnection()
        {
            con.Close();
            con.Dispose();
        }
    }
}
