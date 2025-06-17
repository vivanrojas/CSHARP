using EndesaBusiness.servidores;
using EndesaEntity.medida;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;


namespace EndesaBusiness.medida.pendiente
{
    public class PendienteFacturadoSAPKEE : PendienteSAPKEE
    {
        logs.Log ficheroLog;
        Dictionary<string, List<PendienteSAPKEE>> dic;
        public bool existe { get; set; }

        public PendienteFacturadoSAPKEE()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "");
            dic = Carga();
        }

        private Dictionary<string, List<PendienteSAPKEE>> Carga()
        {

            //////servidores.RedShiftServer db;
            //////OdbcCommand command;
            //////OdbcDataReader r;
            string strSql = "";

            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string CUPS = "";

            Dictionary<string, List<PendienteSAPKEE>> d = new Dictionary<string, List<PendienteSAPKEE>>();

            try
            {

                strSql = "SELECT substring(cd_cups_ext,1,20) as cups, cd_mes as mes, fh_fact,fh_ini_fact, fh_fin_fact "
                + " FROM t_ed_h_sap_facts"
                + " group by substring(cd_cups_ext,1,20) , cd_mes , fh_fact,fh_ini_fact, fh_fin_fact"
                + " ORDER BY substring(cd_cups_ext,1,20), fh_ini_fact desc";

                ficheroLog.Add(strSql);
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    PendienteSAPKEE c = new PendienteSAPKEE();
                    c.cups = r["cups"].ToString();
                    c.mes = r["mes"].ToString();
     
                    c.fecha_factura = Convert.ToDateTime(r["fh_fact"]);
                    if (r["fh_ini_fact"] != System.DBNull.Value) {
                        c.ult_fecha_desde_facturada = Convert.ToDateTime(r["fh_ini_fact"]);
                    }
                    if (r["fh_fin_fact"] != System.DBNull.Value)
                    {
                        c.ult_fecha_hasta_facturada = Convert.ToDateTime(r["fh_fin_fact"]);
                    }

                    if (CUPS != r["cups"].ToString())
                    {
                        List<PendienteSAPKEE> t;
                        if (!d.TryGetValue(c.cups, out t))
                        {
                            t = new List<PendienteSAPKEE>();
                            t.Add(c);
                            d.Add(c.cups, t);
                        }
                        else
                            t.Add(c);


                    }
                    CUPS = r["cups"].ToString();

                }
                db.CloseConnection();
                return d;


            }
            catch (Exception ex)
            {
                ficheroLog.addError("Carga: " + ex.Message);
                return null;
            }
        }




       


        public List<string> GetCupsFinalUltimaFacturada(string cups20)
        {
            List<string> lista = new List<string>();

            List<PendienteSAPKEE> o;
            if (dic.TryGetValue(cups20, out o))
            {

                lista =  o.FindAll(z => z.cups == cups20 ).Select(z => Convert.ToString(z.ult_fecha_hasta_facturada)).ToList();
            }


            return lista;
  
           

        }

        public List<string> GetCupsInicioUltimaFacturada(string cups20)
        {
            List<string> lista = new List<string>();

            List<PendienteSAPKEE> o;
            if (dic.TryGetValue(cups20, out o))
            {

                lista = o.FindAll(z => z.cups == cups20).Select(z => Convert.ToString(z.ult_fecha_desde_facturada)).ToList();
            }


            return lista;



        }


    }
}
