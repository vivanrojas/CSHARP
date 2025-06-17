using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class Kee_Extraccion_Formulas
    {

        public enum Tipo
        {
            CH,
            CCH,
            Todos
        }

        utilidades.Param p;
        public List<EndesaEntity.medida.Kee_Extraccion_Formulas> list { get; set; }
        public List<string> lista_extraccion { get; set; }        

        public Kee_Extraccion_Formulas()
        {
            p = new utilidades.Param("kee_param", MySQLDB.Esquemas.MED);
            list = Load(Tipo.Todos);
            lista_extraccion = CargaListaExtraccion();            

        }

        public Kee_Extraccion_Formulas(Tipo tipo)
        {
            list = Load(tipo);
        }

        private List<EndesaEntity.medida.Kee_Extraccion_Formulas> Load (Tipo tipo)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            List<EndesaEntity.medida.Kee_Extraccion_Formulas> l = new List<EndesaEntity.medida.Kee_Extraccion_Formulas>();
            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            List<string> o;

            try
            {
                //strSql = "SELECT c.cups20,"
                //     + " SUBSTR(c.cups22, 21, 2) sufijo FROM kee_reporte_extraccion_ch c"
                //     + " GROUP BY c.cups20, SUBSTR(c.cups22, 21, 2)";
                //db = new MySQLDB(MySQLDB.Esquemas.MED);
                //command = new MySqlCommand(strSql, db.con);
                //r = command.ExecuteReader();
                //while (r.Read())
                //{
                    
                //    if (!dic.TryGetValue(r["cups20"].ToString(), out o))
                //    {
                //        o = new List<string>();
                //        o.Add(r["sufijo"].ToString());
                //        dic.Add(r["cups20"].ToString(), o);
                //    }
                //    else
                //        o.Add(r["sufijo"].ToString());
                //}

                //db.CloseConnection();

                //strSql = "SELECT c.cups20,"
                //     + " SUBSTR(c.cups22, 21, 2) sufijo FROM kee_reporte_extraccion_cch c"
                //     + " GROUP BY c.cups20, SUBSTR(c.cups22, 21, 2)";
                //db = new MySQLDB(MySQLDB.Esquemas.MED);
                //command = new MySqlCommand(strSql, db.con);
                //r = command.ExecuteReader();
                //while (r.Read())
                //{

                //    if (!dic.TryGetValue(r["cups20"].ToString(), out o))
                //    {
                //        o = new List<string>();
                //        o.Add(r["sufijo"].ToString());
                //        dic.Add(r["cups20"].ToString(), o);
                //    }
                //    else
                //        o.Add(r["sufijo"].ToString());
                //}

                //db.CloseConnection();



                Console.WriteLine("Cargando datos de kee_extraccion_formulas");
                //strSql = "SELECT cups20, fecha_desde, fecha_hasta, tipo, fuente, extraccion, usuario, fecha_archivo, f_ult_mod"
                //    + " FROM med.kee_extraccion_formulas ";
                //switch (tipo)
                //{
                //    case Tipo.CH:
                //        strSql += " where fuente = 'CH'";
                //        break;
                //    case Tipo.CCH:
                //        strSql += " where fuente = 'CCH'";
                //        break;
                //    case Tipo.Todos:
                //        strSql += "";
                //        break;
                //}

                strSql = "SELECT k.cups20,"
                    + " k.fecha_desde AS fecha_sol_desde, k.fecha_hasta AS fecha_sol_hasta,"
                    + " if (k.fuente = 'CH', kd.cups22, kdd.cups22) AS cups22,"
                    + " if (k.fuente = 'CH', kd.fecha_desde, kdd.fecha_desde) AS fecha_desde,"
                    + " if (k.fuente = 'CH', kd.fecha_hasta, kdd.fecha_hasta) AS fecha_hasta,"
                    + " if(k.fuente = 'CH', kd.fuente, kdd.fuente) AS origen,"
                    + " k.fuente AS tipo_curva,"
                    + " k.extraccion"
                    + " FROM kee_extraccion_formulas k"
                    + " LEFT OUTER JOIN kee_reporte_extraccion_ch_r kd ON"
                    + " kd.cups20 = k.cups20 AND"
                    + " kd.fuente = k.tipo"
                    + " LEFT OUTER JOIN kee_reporte_extraccion_cch_r kdd ON"
                    + " kdd.cups20 = k.cups20 AND"
                    + " kdd.fuente = k.tipo";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Kee_Extraccion_Formulas c = new EndesaEntity.medida.Kee_Extraccion_Formulas();
                    c.cups20 = r["cups20"].ToString();

                    if (r["cups22"] != System.DBNull.Value)
                        c.cups22 = r["cups22"].ToString();

                    if (r["fecha_desde"] != System.DBNull.Value)
                        c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);

                    if (r["fecha_hasta"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);

                    c.fecha_sol_desde = Convert.ToDateTime(r["fecha_sol_desde"]);
                    c.fecha_sol_hasta = Convert.ToDateTime(r["fecha_sol_hasta"]);

                    if (r["origen"] != System.DBNull.Value)
                        c.tipo = r["origen"].ToString();

                    c.fuente = r["tipo_curva"].ToString();
                    c.extraccion = r["extraccion"].ToString();


                    //if (r["usuario"] != System.DBNull.Value)
                    //    c.usuario = r["usuario"].ToString();
                    //if (r["fecha_archivo"] != System.DBNull.Value)
                    //    c.fecha_mod_archivo = Convert.ToDateTime(r["fecha_archivo"]);                    

                    l.Add(c);
                }
                db.CloseConnection();
                return l;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public void Subir_FTP()
        {

            EndesaBusiness.utilidades.UltimateFTP ftp;
            FileInfo file;
            EndesaBusiness.medida.PM_ML pmml;

            try
            {
                               

                file = new FileInfo(p.GetValue("output_fichero_CSV") + p.GetValue("archivo_extraccion_formulas").Replace("AAAAMMDD",DateTime.Now.ToString("yyyyMMdd")));

                ftp = new EndesaBusiness.utilidades.UltimateFTP(
                        p.GetValue("ftp_server", DateTime.Now, DateTime.Now),
                        p.GetValue("ftp_user", DateTime.Now, DateTime.Now),
                        p.GetValue("ftp_pass", DateTime.Now, DateTime.Now),
                        p.GetValue("ftp_port", DateTime.Now, DateTime.Now));

                ftp.Upload(p.GetValue("ftp_ruta_extr_form") + file.Name, file.FullName);             


              //  MessageBox.Show("Se ha exportado el archivo "
              //     + archivo + " correctamente.",
              //"Exportación tabla PM ML a FTP",
              //MessageBoxButtons.OK,
              //MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                
             //   MessageBox.Show(e.Message,
             //"Exportación tabla PM ML a FTP",
             //MessageBoxButtons.OK,
             //MessageBoxIcon.Error);
            }
        }

        private List<string> CargaListaExtraccion()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            List<string> l = new List<string>();
            try
            {
                strSql = "SELECT extraccion "
                    + " FROM kee_extraccion_formulas"
                    + " GROUP BY extraccion";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    if (r["extraccion"] != System.DBNull.Value)
                        l.Add(r["extraccion"].ToString());
                }
                db.CloseConnection();
                return l;
            }catch(Exception e)
            {
                return null;
            }
        }

       

        public DateTime MinDateValue()
        {
            return list.Min(z => z.fecha_desde);
        }

        public DateTime MaxDateValue()
        {
            return list.Max(z => z.fecha_hasta);
        }


    }
}
