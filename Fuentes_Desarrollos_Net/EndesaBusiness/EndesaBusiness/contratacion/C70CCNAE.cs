using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion
{
    public class C70CCNAE
    {
        logs.Log ficheroLog;
        utilidades.Param param;

        // Para el control de la ejecucion
        EndesaBusiness.utilidades.Seguimiento_Procesos ss_pp;
        public C70CCNAE()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_C70CNAE");
            param = new utilidades.Param("c70ccnae_param", servidores.MySQLDB.Esquemas.CON);
            ss_pp = new utilidades.Seguimiento_Procesos();
        }
            
        public void DescargaC70CNAE()
        {
            string extractor = "";
            string archivo_cnae = "";
            string md5 = "";
            utilidades.Fechas utilFechas = new utilidades.Fechas();

            try
            {
                extractor = param.GetValue("extractor_c70cnae");

                ss_pp.Update_Fecha_Inicio("Contratación", "C70CNAE", "1_Ejecución Extractor");

                ficheroLog.Add("Ejecutando extractor: " + extractor);                
                utilidades.Fichero.EjecutaComando(extractor);
                ficheroLog.Add("Extractor finalizado.");

                ss_pp.Update_Fecha_Fin("Contratación", "C70CNAE", "1_Ejecución Extractor");

                archivo_cnae = param.GetValue("inbox") + param.GetValue("nombre_archivo_extractor");
                //archivo_cnae = archivo_cnae.Replace("YYMMDD", utilFechas.UltimoDiaHabil().ToString("yyyyMMdd"));
                archivo_cnae = archivo_cnae.Replace("YYYYMMDD", DateTime.Now.ToString("yyyyMMdd"));
                md5 = utilidades.Fichero.checkMD5(archivo_cnae).ToString();

                if (md5 != param.GetValue("md5_fichero"))
                {
                    ss_pp.Update_Fecha_Inicio("Contratación", "C70CNAE", "2_Importación");
                    // ************************
                    Carga_C70CCNAE(archivo_cnae);
                    // ************************
                    ss_pp.Update_Fecha_Fin("Contratación", "C70CNAE", "2_Importación");

                    param.code = "md5_fichero";
                    param.from_date = new DateTime(2022, 05, 11);
                    param.to_date = new DateTime(2099, 12, 31);
                    param.value = md5;
                    param.Save();

                }
                else
                    ss_pp.Update_Comentario("Contratación", "C70CNAE", "2_Importación",
                        "El archivo " + param.GetValue("nombre_archivo_extractor")
                        + " no se ha actualizado.");

            }
            catch(Exception ex)
            {
                ss_pp.Update_Comentario("Contratación", "C70CNAE", "1_Ejecución Extractor", ex.Message);
                ficheroLog.AddError("DescargaC70CNAE: " + ex.Message);
            }
        }



        public void Carga_C70CCNAE(String nombre_archivo)
        {
            long i = 0;
            string line;
            StringBuilder sb = new StringBuilder();
            Boolean firstOnly = true;
            MySQLDB db;
            MySqlCommand command;

            try
            {

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand("delete from cont.c70ccnae;", db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                System.IO.StreamReader file = new System.IO.StreamReader(nombre_archivo);
                while ((line = file.ReadLine()) != null)
                {

                    if (firstOnly)
                    {
                        sb.Append("REPLACE INTO c70ccnae (");
                        sb.Append("cups13,cnae,complemento) values ");
                        firstOnly = false;
                    }

                    if (line.Length > 50)
                    {
                        string[] f = line.Split(';');
                        i++;
                        sb.Append("('").Append(f[4]).Append("',"); // cups13
                        sb.Append("'").Append(f[6].Trim()).Append("',"); // cnae
                        sb.Append("'").Append(f[13].Trim()).Append("'),"); // complemento                        
                    }


                    if (i == 250)
                    {
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.CON);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        i = 0;
                    }
                }

                if (i > 0)
                {
                    firstOnly = true;
                    db = new MySQLDB(MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                }

                file.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                  "Error en la importación de la extracción C70CCNAE.txt",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }
    }
}
