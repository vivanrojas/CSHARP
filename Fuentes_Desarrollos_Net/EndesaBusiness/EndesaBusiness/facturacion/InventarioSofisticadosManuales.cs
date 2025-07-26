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
    public class InventarioSofisticadosManuales : EndesaEntity.facturacion.PuntosSofisticados
    {

        public List<EndesaEntity.facturacion.PuntosSofisticados> lista { get; set; }
        utilidades.Global g = new utilidades.Global();
        public InventarioSofisticadosManuales()
        {

        }

        public void Carga(string cups13)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime begin = new DateTime();

            try
            {
                lista = new List<EndesaEntity.facturacion.PuntosSofisticados>();
                begin = DateTime.Now;

                strSql = "Select CCOUNIPS, CUPS20, DAPERSOC, GRUPO, FD, FH, PRECIOS, FACTURAS_A_CUENTA"
                 + " from fact.cm_sofisticados where CCOUNIPS <> ''";

                if (cups13 != null)
                {
                    strSql += " and CCOUNIPS = '" + cups13 + "'";
                }

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.facturacion.PuntosSofisticados c = new EndesaEntity.facturacion.PuntosSofisticados();
                    if (r["CCOUNIPS"] != System.DBNull.Value)
                        c.cups13 = r["CCOUNIPS"].ToString();
                    if (r["CUPS20"] != System.DBNull.Value)
                        c.cups20 = r["CUPS20"].ToString();
                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.dapersoc = r["DAPERSOC"].ToString();
                    if (r["GRUPO"] != System.DBNull.Value)
                        c.grupo = r["GRUPO"].ToString();
                    if (r["FD"] != System.DBNull.Value)
                        c.fd = Convert.ToDateTime(r["FD"]);
                    if (r["FH"] != System.DBNull.Value)
                        c.fh = Convert.ToDateTime(r["FH"]);
                    if (r["PRECIOS"] != System.DBNull.Value)
                        c.precios = r["PRECIOS"].ToString();
                    if (r["FACTURAS_A_CUENTA"] != System.DBNull.Value)
                        c.facturas_a_cuenta = r["FACTURAS_A_CUENTA"].ToString() == "S";

                    lista.Add(c);
                }

                if (lista.Count == 0)
                {
                    MessageBox.Show("La consulta que ha realizado no devuelve ningún resultado.",
                    "La consulta no devuelte datos.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                }

                Cursor.Current = Cursors.Default;

                g.SaveQuery("FrmPuntosSofisticados", strSql, begin, DateTime.Now);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "Error en la búsqueda de Puntos Sofisticados",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            }
        }

        public void Add()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            

            try
            {
                strSql = "replace into cm_sofisticados (CCOUNIPS, CUPS20, DAPERSOC, GRUPO, FD, FH, PRECIOS, FACTURAS_A_CUENTA,"
                    + " USER) values"
                    + " ('" + cups13 + "',"
                    + (cups20 != null ? "'" + cups20 + "'," : "null,")
                    + (dapersoc != null ? "'" + dapersoc + "'," : "null,")
                    + (grupo != null ? "'" + grupo + "'," : "null,")
                    + "'" + fd.ToString("yyyy-MM-dd") + "',"
                    + (fh != null ? "'" + fh.ToString("yyyy-MM-dd") + "'," : "null,")
                    + (precios != null ? "'" + precios + "'," : "null,")
                    + "'" + (facturas_a_cuenta ? "S" : "N") + "','" + Environment.UserName + "');";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormPuntosSofisticados", strSql, DateTime.Now, DateTime.Now);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el guardado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public void Del()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            


            try
            {
                strSql = "delete from cm_sofisticados where codigo = '" + cups13 + "'"
                    + " and FD = '" + fd.ToString("yyyy-MM-dd");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormPuntosSofisticados", strSql, DateTime.Now, DateTime.Now);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el borrado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        public void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            

            try
            {
                strSql = "update cm_sofisticados set facturas_a_cuenta = '" + (facturas_a_cuenta ? "S" : "N") + "'";

                if (precios != null)
                {
                    strSql = strSql + " ,PRECIOS = '" + precios + "'";
                }

                if (dapersoc != null)
                {
                    strSql = strSql + " ,DAPERSOC = '" + dapersoc + "'";
                }
                if (grupo != null)
                {
                    strSql = strSql + " ,GRUPO = '" + grupo + "'";
                }
                if (fh != null)
                {
                    strSql = strSql + " ,FH = '" + fh.ToString("yyyy-MM-dd") + "'";
                }


                strSql = strSql + " ,user = '" + Environment.UserName + "'";

                strSql = strSql + " where CCOUNIPS = '" + cups13 + "'"
                    + " and FD = '" + fd.ToString("yyyy-MM-dd") + "'";


                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                g.SaveQuery("FormPuntosSofisticados", strSql, DateTime.Now, DateTime.Now);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el guardado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }
    }
}
