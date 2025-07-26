using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{
    public class SofisticadosProgramasConsumo
    {
        public Dictionary<string, EndesaEntity.medida.DiccionarioCurva> dic { get; set; }

        public SofisticadosProgramasConsumo(string cliente, DateTime fd, DateTime fh)
        {
            dic = new Dictionary<string, EndesaEntity.medida.DiccionarioCurva>();
            Carga(cliente, fd, fh);
        }

        public void Carga(string cliente, DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            int numPeriodos = 0;
            bool firstOnly = true;

            DateTime fechaHora = new DateTime();
            utilidades.Fechas utilFechas = new utilidades.Fechas();

            try
            {
                strSql = "SELECT i.cups20, p.Fecha";
                for (int i = 1; i <= 25; i++)
                    strSql += ",Value" + i;

                strSql += " FROM ag_inventario i INNER JOIN"
                    + " ag_programas p on"
                    + " p.CCOUNIPS = i.cups13"
                    + " where i.cliente = '" + cliente + "' and"
                    + " (p.Fecha >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " p.Fecha <= '" + fh.ToString("yyyy-MM-dd") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    fechaHora = Convert.ToDateTime(r["Fecha"]);
                    numPeriodos = utilFechas.NumHoras(fechaHora, fechaHora);
                    firstOnly = true;

                    for (int i = 1; i <= 25; i++)
                    {
                        if (firstOnly)
                        {
                            fechaHora = Convert.ToDateTime(r["Fecha"]);
                            firstOnly = false;
                        }
                        else
                        {
                            switch (numPeriodos)
                            {
                                case 24:
                                    fechaHora = fechaHora.AddHours(1);
                                    break;
                            }
                        }


                        if (r["Value" + i] != System.DBNull.Value)
                        {
                            EndesaEntity.medida.Curva c = new EndesaEntity.medida.Curva();
                            c.a = Convert.ToDouble(r["Value" + i]);
                            c.origen = "P";
                            c.r = 0;


                            EndesaEntity.medida.DiccionarioCurva o;
                            if (!dic.TryGetValue(r["cups20"].ToString(), out o))
                            {
                                o = new EndesaEntity.medida.DiccionarioCurva();

                                EndesaEntity.medida.Curva oo;
                                if (!o.dic.TryGetValue(fechaHora, out oo))
                                {
                                    o.dic.Add(fechaHora, c);
                                    dic.Add(r["cups20"].ToString(), o);
                                }
                            }
                            else
                            {
                                EndesaEntity.medida.Curva oo;
                                if (!o.dic.TryGetValue(fechaHora, out oo))
                                    o.dic.Add(fechaHora, c);
                            }

                        }

                    }

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "ProgrmasConsumo.Carga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public void GuardaCurva(Dictionary<string, EndesaEntity.medida.DiccionarioCurva> dic)
        {
            MySQLDB db;
            MySqlCommand command;
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            utilidades.Fechas utilFechas = new utilidades.Fechas();
            int totalRegistros = 0;


            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand("replace into ag_ch_hist select * from ag_ch", db.con);
            command.ExecuteReader();
            db.CloseConnection();

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand("delete from ag_ch", db.con);
            command.ExecuteReader();
            db.CloseConnection();


            foreach (KeyValuePair<string, EndesaEntity.medida.DiccionarioCurva> p in dic)
            {
                foreach (KeyValuePair<DateTime, EndesaEntity.medida.Curva> pp in p.Value.dic)
                {

                    if (firstOnly)
                    {
                        sb.Append("replace into ag_ch (cups20, fecha, hora, estacion, activa, reactiva, origen, usuario, f_ult_mod) values ");
                        firstOnly = false;
                    }

                    totalRegistros++;
                    sb.Append("('").Append(p.Key).Append("',");
                    sb.Append("'").Append(pp.Key.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append(pp.Key.Hour).Append(",");
                    sb.Append(utilFechas.Estacion(pp.Key)).Append(",");
                    if (pp.Value == null)
                        sb.Append("null, null, null,  ");
                    else
                    {
                        sb.Append(pp.Value.a.ToString().Replace(",", ".")).Append(",");
                        sb.Append(pp.Value.r.ToString().Replace(",", ".")).Append(",");
                        sb.Append("'").Append(pp.Value.origen).Append("',");
                    }

                    sb.Append("'").Append(System.Environment.UserName).Append("',");
                    sb.Append("'").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm")).Append("'),");

                    if (totalRegistros == 250)
                    {
                        totalRegistros = 0;
                        firstOnly = true;
                        db = new MySQLDB(MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteReader();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                    }

                }


            }

            if (totalRegistros > 0)
            {
                firstOnly = true;
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                command.ExecuteReader();
                db.CloseConnection();
                sb = null;

            }
        }
    }
}
