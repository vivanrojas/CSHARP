using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.cuadro_mando
{
    public class TopPortugal_Electricidad
    {
        public Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> dic;
        public TopPortugal_Electricidad()
        {
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> Carga()
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe> d =
                new Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Informe>();

            int maxMes = 0;

            try
            {

                // Se cambia de una lista fija de puntos a 
                // desde el inventario de puntos de MT con TAM > 100.000

                //strSql = "select max(aaaamm) MaxMes from cm_toppt";
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //r = command.ExecuteReader();
                //while (r.Read())
                //{
                //    maxMes = Convert.ToInt32(r["MaxMes"]);
                //}
                //db.CloseConnection();

                //strSql = "SELECT CNIFDNIC, DAPERSOC, CCOUNIPS, CUPSREE, aaaamm"
                //    + " FROM fact.cm_toppt where"
                //    + " aaaamm = " + maxMes;
                //db = new MySQLDB(MySQLDB.Esquemas.FAC);
                //command = new MySqlCommand(strSql, db.con);
                //r = command.ExecuteReader();

                //GUS - 13/09/2024 - Deshabilitamos la carga de puntos Portugal MT de SCE

                ////strSql = "SELECT mt.CNIFDNIC, mt.DAPERSOC, mt.CUPS as CCOUNIPS, mt.CUPSREE"
                ////    + " FROM cont.contratos_ps_mt mt"
                ////    + " left outer join fact.tam t ON"
                ////    + " t.CNIFDNIC = mt.CNIFDNIC and"
                ////    + " t.CUPS20 = mt.CUPSREE"
                ////    + " WHERE t.TAM > 100000";
                ////db = new MySQLDB(MySQLDB.Esquemas.GBL);
                ////command = new MySqlCommand(strSql, db.con);
                ////r = command.ExecuteReader();
                ////while (r.Read())
                ////{
                ////    EndesaEntity.facturacion.cuadroDeMando.Informe c = new EndesaEntity.facturacion.cuadroDeMando.Informe();
                ////    c.origen_sistemas = "SCE";
                ////    c.nif = r["CNIFDNIC"].ToString();
                ////    c.cliente = r["DAPERSOC"].ToString();
                ////    c.cups13 = r["CCOUNIPS"].ToString();
                ////    c.cups20 = r["CUPSREE"].ToString();

                ////    EndesaEntity.facturacion.cuadroDeMando.Informe o;
                ////    if (!d.TryGetValue(c.cups20, out o))
                ////        d.Add(c.cups20, c);

                ////}
                ////db.CloseConnection();

                strSql = "SELECT pt.cd_nif_cif_cli AS CNIFDNIC,"
                    + " pt.tx_apell_cli AS DAPERSOC, pt.cd_cups AS CCOUNIPS,"
                    + " pt.cups20 AS CUPSREE,"
                    + " a.agora, a.tam"
                    + " FROM cont.t_ed_h_ps_pt pt"
                    + " INNER JOIN fact.t_ed_h_sap_tam_agora a ON"
                    + " a.cd_cups = pt.cups20"
                    + " WHERE pt.cd_tp_tension = 'MT' and"
                    + " a.tam > 100000";
                db = new MySQLDB(MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.cuadroDeMando.Informe c = new EndesaEntity.facturacion.cuadroDeMando.Informe();
                    c.origen_sistemas = "SAP";
                    c.nif = r["CNIFDNIC"].ToString();
                    c.cliente = r["DAPERSOC"].ToString();
                    c.cups13 = r["CCOUNIPS"].ToString();
                    c.cups20 = r["CUPSREE"].ToString();

                    EndesaEntity.facturacion.cuadroDeMando.Informe o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();


                return d;

            }catch(Exception ex)
            {
                return null;
            }
        }


    }
}
