using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EndesaBusiness.medida
{
    public class CurvaResumenEER
    {

        Dictionary<string, List<EndesaEntity.medida.CurvaResumen>> dic;
        public int version_curva { get; set; }        

        public CurvaResumenEER(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }


        private Dictionary<string, List<EndesaEntity.medida.CurvaResumen>> Carga(DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string myKey = "";

            Dictionary<string, List<EndesaEntity.medida.CurvaResumen>> d =
                new Dictionary<string, List<EndesaEntity.medida.CurvaResumen>>();

            try
            {
                strSql = "SELECT cups20, tarifa, fd, fh, version, estado, num_dias, activa, reactiva,"
                    + " a_p1, a_p2, a_p3, a_p4, a_p5, a_p6,"
                    + " r_p1, r_p2, r_p3, r_p4, r_p5, r_p6,"
                    + " potmax_p1, potmax_p2, potmax_p3, potmax_p4, potmax_p5, potmax_p6,"
                    + " origen, completa, num_periodos, f_ult_mod"
                    + " FROM cont.eer_resumen_medida"
                    + " where (fd >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fh <= '" + fh.ToString("yyyy-MM-dd") + "')"
                    + " order by cups20, fd, version";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.CurvaResumen c = new EndesaEntity.medida.CurvaResumen();

                    c.cups20 = r["cups20"].ToString();
                    c.tarifa = r["tarifa"].ToString();
                    c.fd = Convert.ToDateTime(r["fd"]);
                    c.fh = Convert.ToDateTime(r["fh"]);
                    c.version = Convert.ToInt32(r["version"]);
                    c.estado = r["estado"].ToString();
                    c.activa = Convert.ToInt32(r["activa"]);
                    c.reactiva = Convert.ToInt32(r["reactiva"]);

                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["a_p" + i] != System.DBNull.Value)
                            c.activa_periodo[i] = Convert.ToDouble(r["a_p" + i]);
                        if (r["r_p" + i] != System.DBNull.Value)
                            c.reactiva_periodo[i] = Convert.ToDouble(r["r_p" + i]);
                        if (r["potmax_p" + i] != System.DBNull.Value)
                            c.potencias_maximas[i] = Convert.ToDouble(r["potmax_p" + i]);
                    }

                    c.origen = r["origen"].ToString();
                    c.completa = r["completa"].ToString() == "S";
                    c.num_periodos = Convert.ToInt32(r["num_periodos"]);

                    myKey = c.cups20 + "_"
                        + c.fd.ToString("yyyyMMdd") + "_"
                        + c.fh.ToString("yyyyMMdd");

                    List<EndesaEntity.medida.CurvaResumen> o;
                    if (!d.TryGetValue(myKey, out o))                    
                        o = new List<EndesaEntity.medida.CurvaResumen>();

                    o.Add(c);

                    d.Add(myKey, o);

                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public void Save(EndesaEntity.medida.CurvaResumen cr)
        {
            string myKey = "";            
            
            myKey = cr.cups20 + "_"
                + cr.fd.ToString("yyyyMMdd") + "_"
                + cr.fh.ToString("yyyyMMdd");

            List<EndesaEntity.medida.CurvaResumen> o;
            if (dic.TryGetValue(myKey, out o))
            {

                // Comprobamos que la última curva resumen es distinta a la que estamos
                // tratando.

                EndesaEntity.medida.CurvaResumen c = o[o.Count() - 1];
                this.version_curva = c.version;
                //if (c.completa && (c.activa != cr.activa || c.reactiva != cr.reactiva) && cr.completa)
                //{
                //    cr.version = c.version + 1;
                //    this.version_curva = cr.version;
                //    cr.estado = "R";
                //    SaveMySQL(cr);
                //}
                if(c.estado != "F" && cr.completa)
                {
                    cr.version = 1;
                    this.version_curva = cr.version;
                    cr.estado = "R";
                    SaveMySQL(cr);
                }                   
               
            }
            else
            {
                cr.version = 1;
                version_curva = cr.version;
                cr.estado = "R";
                SaveMySQL(cr);
            }



        }

        public bool CurvaCompleta(string cups20, DateTime fd, DateTime fh)
        {
            bool completa = false;

            if(dic != null)
            {
                List<EndesaEntity.medida.CurvaResumen> o;
                if (dic.TryGetValue(GetKey(cups20, fd, fh), out o))
                {
                    EndesaEntity.medida.CurvaResumen p = o.FindLast(z => z.fd == fd && z.fh == fh);
                    if (p != null)
                        return p.completa;
                }
            }           

            return completa;
        }
        
        private void SaveMySQL(EndesaEntity.medida.CurvaResumen cr)
        {
            StringBuilder sb = new StringBuilder();
            MySQLDB db;
            MySqlCommand command;
            try
            {
                sb.Append("replace into eer_resumen_medida ");
                sb.Append("(cups20, tarifa, fd, fh, version, estado, num_dias, activa, reactiva,");
                sb.Append(" a_p1, a_p2, a_p3, a_p4, a_p5, a_p6, r_p1, r_p2, r_p3, r_p4, r_p5, r_p6,");
                sb.Append(" potmax_p1, potmax_p2, potmax_p3, potmax_p4, potmax_p5, potmax_p6, origen,");
                sb.Append(" completa, num_periodos) values ");

                sb.Append("('").Append(cr.cups20).Append("',");
                sb.Append("'").Append(cr.tarifa).Append("',");
                sb.Append("'").Append(cr.fd.ToString("yyyy-MM-dd")).Append("',");
                sb.Append("'").Append(cr.fh.ToString("yyyy-MM-dd")).Append("',");
                sb.Append(cr.version).Append(",");
                sb.Append("'").Append(cr.estado).Append("',");
                sb.Append(Convert.ToInt32((cr.fh - cr.fd).TotalDays + 1)).Append(",");
                sb.Append(cr.activa.ToString().Replace(",",".")).Append(",");
                sb.Append(cr.reactiva.ToString().Replace(",", ".")).Append(",");

                for(int i = 1; i <= 6; i++)
                {
                    if (cr.activa_periodo[i] != 0)
                        sb.Append(cr.activa_periodo[i].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                }

                for (int i = 1; i <= 6; i++)
                {
                    if (cr.reactiva_periodo[i] != 0)
                        sb.Append(cr.reactiva_periodo[i].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                }

                for (int i = 1; i <= 6; i++)
                {
                    if (cr.potencias_maximas[i] != 0)
                        sb.Append(cr.potencias_maximas[i].ToString().Replace(",", ".")).Append(",");
                    else
                        sb.Append("null,");
                }

                sb.Append("'").Append(cr.origen).Append("',");
                if (cr.completa)
                    sb.Append("'S',");
                else
                    sb.Append("'N',");

                sb.Append(cr.num_periodos).Append(");");

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(sb.ToString(), db.con);
                command.ExecuteReader();
                db.CloseConnection();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private string GetKey(string cups20, DateTime fd, DateTime fh)
        {
            string myKey = cups20 + "_"
                + fd.ToString("yyyyMMdd") + "_"
                + fh.ToString("yyyyMMdd");

            return myKey;
        }

        public EndesaEntity.medida.CurvaResumen GetCurvaResumen(string cups, DateTime fd, DateTime fh)
        {

            string clave = "";

            clave = cups + "_" + fd.ToString("yyyyMMdd") + "_" + fh.ToString("yyyyMMdd");

            EndesaEntity.medida.CurvaResumen r;
            List<EndesaEntity.medida.CurvaResumen> o;

            if (dic.TryGetValue(clave, out o))
            {
                r = o.Find(z => z.fd == fd && z.fh == fh);
                return r;
            }
            else
                return null;
            
        }

    }
}
