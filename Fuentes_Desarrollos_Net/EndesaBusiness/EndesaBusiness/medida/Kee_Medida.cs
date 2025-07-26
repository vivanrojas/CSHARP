using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.medida
{
    public class Kee_Medida
    {
        public Dictionary<string, EndesaEntity.medida.DiccionarioCurva> dic { get; set; }
        public Kee_Medida()
        {
            dic = new Dictionary<string, EndesaEntity.medida.DiccionarioCurva>();
        }

        public Kee_Medida(List<string> lista_cups22, string origen, DateTime fd, DateTime fh)
        {
            dic = new Dictionary<string, EndesaEntity.medida.DiccionarioCurva>();
            Carga(lista_cups22, origen, fd, fh);
        }

        private void Carga(List<string> lista_cups22, string origen, DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            DateTime fechaHora = new DateTime();

            try
            {
                strSql = "SELECT ch.cups22, ch.fecha, ch.hora, ch.ae, ch.r1 FROM kee_reporte_extraccion_ch ch"
                    + " WHERE ch.cups22 in ('" + lista_cups22[0] + "'";

                for (int i = 1; i < lista_cups22.Count; i++)
                    strSql += ",'" + lista_cups22[i] + "'";

                strSql += ") AND ch.fuente = '" + origen + "'"
                    + " AND (ch.fecha >= '" + fd.ToString("yyyy-MM-dd") + "'"
                    + " AND ch.fecha <= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (r["ae"] != System.DBNull.Value)
                    {
                        EndesaEntity.medida.Curva c = new EndesaEntity.medida.Curva();
                        c.a = Convert.ToInt32(r["ae"]);
                        c.origen = origen;

                        if (r["r1"] != System.DBNull.Value)
                            c.r = Convert.ToInt32(r["r1"]);

                        fechaHora = Convert.ToDateTime(r["fecha"].ToString());
                        fechaHora = fechaHora.AddHours(Convert.ToInt32(r["hora"].ToString().Substring(0, 2)));

                        EndesaEntity.medida.DiccionarioCurva o;
                        if (!dic.TryGetValue(r["cups22"].ToString(), out o))
                        {
                            o = new EndesaEntity.medida.DiccionarioCurva();
                            o.dic.Add(fechaHora, c);
                            dic.Add(r["cups22"].ToString(), o);
                        }
                        else
                            o.dic.Add(fechaHora, c);

                    }

                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Kee_Medida.Carga",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }


        }
    }
}
