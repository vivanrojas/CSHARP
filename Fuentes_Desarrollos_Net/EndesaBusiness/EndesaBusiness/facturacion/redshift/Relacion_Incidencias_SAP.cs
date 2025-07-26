using EndesaBusiness.servidores;
using EndesaBusiness.calendarios;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace EndesaEntity.medida
{
    public class Relacion_Incidencias_SAP : RelacionIncidencias
    {
        ////public Dictionary<string, RelacionIncidencias> dic { get; set; }

        Dictionary<string, List<RelacionIncidencias>> dic;
        public bool existe { get; set; }

        public Relacion_Incidencias_SAP()
        {
            dic = Carga();
        }

        ////private Dictionary<string, RelacionIncidencias> Carga()
        private Dictionary<string, List<RelacionIncidencias>> Carga()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string cadenacontrol;

            ////Dictionary<string, RelacionIncidencias> d = new Dictionary<string, RelacionIncidencias>();
            Dictionary<string, List<RelacionIncidencias>> d = new Dictionary<string, List<RelacionIncidencias>>();
            try
            {
                strSql = "SELECT incidencia,cups, periodo_pendiente, Estado_fac_SE, Titulo_Fac FROM Relacion_Incidencias_SAP ORDER BY cups asc, periodo_pendiente asc";
                cadenacontrol = "";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    RelacionIncidencias c = new RelacionIncidencias();

                    c.Incidencia = r["incidencia"].ToString();
                    c.cups = r["cups"].ToString();
                    c.periodo = r["periodo_pendiente"].ToString();
                    c.Estado_Fac_SE = r["Estado_fac_SE"].ToString();
                    c.Titulo_FAC = r["Titulo_Fac"].ToString();

                    Debug.WriteLine(c.cups);
                    //////if (cadenacontrol != c.cups + c.periodo)
                    //////{
                    //////    d.Add(c.cups + c.periodo, c);
                    //////    cadenacontrol = c.cups + c.periodo;
                    //////}

                    List<RelacionIncidencias> t;
                    if (cadenacontrol != c.cups + c.periodo)
                    {
                        cadenacontrol = c.cups + c.periodo;
                        t = new List<RelacionIncidencias>();
                        t.Add(c); 
                        d.Add(cadenacontrol, t);     
                    }
                    

                }
                db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public bool Cups(string cupsmes)
        {
            List<RelacionIncidencias> o;
            if (dic.TryGetValue(cupsmes, out o))
                return true;
            else
                return false;
        }


        public List<string> GetCups(string cupsmes)
        {
            List<string> lista = new List<string>();
            List<RelacionIncidencias> o;

            if (dic.TryGetValue(cupsmes, out o))
            {
                lista = o.FindAll(z => z.cups + z.periodo == cupsmes).Select(z => Convert.ToString(z.Incidencia) + ";" + Convert.ToString(z.Estado_Fac_SE) + ";" + Convert.ToString(z.Titulo_FAC)).ToList();
            }
            return lista;
        }


    }
}