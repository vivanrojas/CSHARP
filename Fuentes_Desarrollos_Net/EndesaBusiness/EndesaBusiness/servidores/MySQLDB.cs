using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.servidores
{
    public class MySQLDB
    {
        private Dictionary<string, string> lista_parametros;
        MySqlTransaction transaction;
        public MySqlConnection con;

        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MySQLDB");

        public enum Esquemas
        {
            MED,
            FAC,
            COB,
            AUX,
            CON,
            GBL,
            Holding,
            HoldingRead,
            Iberdrxxi,
            inffact,
            RAM_PAGOS_POR,
            EER

        }

        public MySQLDB(Esquemas esquema)
        {
            try
            {
                lista_parametros = new Dictionary<string, string>();
                CargaParametros();
                con = new MySqlConnection(GetConnectionString(esquema));
                con.Open();
                
            }
            catch (MySqlException e)
            {
                ficheroLog.AddError("MySQLDB: " + e.Message);
                
            }

        }

        public void MySQLTransaction()
        {
            transaction = con.BeginTransaction();
        }
        public void MySQLCommit()
        {
            transaction.Commit();
        }

        private void CargaParametros()
        {
            string line;
            int pos;
            string key;
            string value;

            FileInfo file = new FileInfo(System.Environment.CurrentDirectory + @"\bin\" + "properties");
            if (!file.Exists)
                MessageBox.Show("No se encuentra el archivo " + System.Environment.CurrentDirectory + @"\bin\" + "properties",
                    "MySQLDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                System.IO.StreamReader archivo = new System.IO.StreamReader(File.OpenRead(file.FullName));
                while ((line = archivo.ReadLine()) != null)
                {
                    if (line.Trim() != "")
                    {
                        pos = line.IndexOf('=');
                        key = line.Substring(0, pos);
                        value = line.Substring(pos + 1);
                        string a;
                        if (!lista_parametros.TryGetValue(key, out a))
                            lista_parametros.Add(key, value);
                    }

                }
                archivo.Close();
            }
        }

        public void CloseConnection()
        {
            con.Close();
            con.Dispose();
            // Console.WriteLine("Se ha desconectado");
        }

        private string GetConnectionString(Esquemas esquema)
        {
            string server = null;
            string ddbb = null;
            string user = null;
            string pass = null;

            // **********************************************
            // **********************************************



            // **********************************************
            // **********************************************



            server = GetValue("server_mysql_siope");
            switch (esquema)
            {
                case Esquemas.MED:
                    ddbb = GetValue("ddbb_med");
                    user = GetValue("ddbb_med_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_med_pass"), true);
                    break;
                case Esquemas.FAC:
                    ddbb = GetValue("ddbb_fac");
                    user = GetValue("ddbb_fac_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_fac_pass"), true);
                    break;
                case Esquemas.COB:
                    ddbb = GetValue("ddbb_cob");
                    user = GetValue("ddbb_cob_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_cob_pass"), true);
                    break;
                case Esquemas.CON:
                    ddbb = GetValue("ddbb_con");
                    user = GetValue("ddbb_con_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_con_pass"), true);
                    break;
                case Esquemas.AUX:
                    ddbb = GetValue("ddbb_aux");
                    user = GetValue("ddbb_aux_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_aux_pass"), true);
                    break;
                case Esquemas.GBL:
                    ddbb = "";
                    user = GetValue("ddbb_glb_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_glb_pass"), true);
                    break;
                case Esquemas.Holding:
                    server = GetValue("server_mysql_pam");
                    ddbb = GetValue("ddbb_hld");
                    user = GetValue("ddbb_hld_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_hld_pass"), true);
                    break;
                case Esquemas.HoldingRead:
                    server = GetValue("server_mysql_pam");
                    ddbb = GetValue("ddbb_hld_read");
                    user = GetValue("ddbb_hld_read_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_hld_read_pass"), true);
                    break;
                case Esquemas.Iberdrxxi:
                    server = GetValue("server_mysql_pam");
                    ddbb = GetValue("ddbb_iber");
                    user = GetValue("ddbb_iber_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_iber_pass"), true);
                    break;                    
                case Esquemas.RAM_PAGOS_POR:
                    server = GetValue("server_mysql_pam");
                    ddbb = GetValue("ddbb_pagos_por");
                    user = GetValue("ddbb_pagos_por_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_pagos_por_pass"), true);
                    break;
                case Esquemas.EER:                    
                    ddbb = GetValue("ddbb_eer");
                    user = GetValue("ddbb_eer_user");
                    pass = utilidades.FuncionesTexto.Decrypt(GetValue("ddbb_eer_pass"), true);
                    break;
                default:
                    MessageBox.Show("No se ha encontrado esquema", "MySQLDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }

            return @"server=" + server + ";userid=" + user + ";password=" + pass + ";database=" + ddbb 
                    + ";Allow User Variables=true;allowLoadLocalInfile=true;UseCompression=true;"
                    + "default command timeout=0;SslMode=None"; 

        }


        private void CargaListaServidores()
        {

        }

        private string GetValue(string key)
        {
            return lista_parametros.First(z => z.Key == key).Value;
        }
    }
}
