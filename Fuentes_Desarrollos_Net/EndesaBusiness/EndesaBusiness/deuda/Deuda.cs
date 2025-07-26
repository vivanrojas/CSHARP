using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.deuda
{
    class Deuda : EndesaEntity.cobros.Deuda_Tabla
    {
        public Dictionary<string, List<EndesaEntity.cobros.Deuda_Tabla>> dic { get; set; }
        public List<EndesaEntity.cobros.Deuda_Tabla> lista { get; set; }
        public Deuda(List<string> lista_nifs)
        {
            dic = new Dictionary<string, List<EndesaEntity.cobros.Deuda_Tabla>>();
            Carga(lista_nifs);
        }


        public void Carga(List<string> lista_nifs)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            bool firstOnly = true;

            try
            {

                strSql = "SELECT dd.CNIFDNIC, dd.DAPERSOC, ps.CUPS22, dd.CFACTURA, dd.IOBLIGAC, dd.FLIMPAGO, dd.FFACTURA,"
                    + " dd.FFACTDES, dd.FFACTHAS, tf.descripcion as TFACTURA,"
                    + " if (dd.TESTIMPG = '016', 'No',"
                        + " if (dd.INTIEXP <> '','No',"
                            + " if (dd.CMEDASJU = '0000','No','Si'))) AS AAJJ,"
                    + " if (TESTIMPG = '016','No',"
                        + " if (CMEDAGER = '0000','No','Si')) AS AGRECOBRO,"
                    + " if (TESTIMPG = '016','No',"
                        + " if (INTIEXP = '','No','Si')) AS PC,"
                    + " if (TESTIMPG = '016','Si','No') AS FC"
                    + " FROM fact.deuda_obligaciones_original dd"
                    + " INNER JOIN cont.PS_AT ps ON"
                    + " ps.IDU = dd.CCOUNIPS"
                    + " LEFT OUTER JOIN fact.fo_p_tipos_factura tf ON"
                    + " tf.id_tipo_factura = dd.TFACTURA"
                    + " WHERE dd.CNIFDNIC in (";

                for (int i = 0; i < lista_nifs.Count; i++)
                    if (firstOnly)
                    {
                        strSql += "'" + lista_nifs[i] + "'";
                        firstOnly = false;
                    }
                    else
                        strSql += " ,'" + lista_nifs[i] + "'";
                strSql += ");";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.GBL);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.cobros.Deuda_Tabla c = new EndesaEntity.cobros.Deuda_Tabla();

                    if (r["CNIFDNIC"] != System.DBNull.Value)
                        c.nif = r["CNIFDNIC"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.dapersoc = r["DAPERSOC"].ToString();

                    if (r["DAPERSOC"] != System.DBNull.Value)
                        c.dapersoc = r["DAPERSOC"].ToString();

                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();

                    if (r["CFACTURA"] != System.DBNull.Value)
                        c.cfactura = r["CFACTURA"].ToString();

                    if (r["IOBLIGAC"] != System.DBNull.Value)
                        c.importe_obligacion = Convert.ToDouble(r["IOBLIGAC"]);

                    if (r["FLIMPAGO"] != System.DBNull.Value)
                        c.fecha_limite_pago = Convert.ToDateTime(r["FLIMPAGO"]);

                    if (r["FFACTURA"] != System.DBNull.Value)
                        c.fecha_factura = Convert.ToDateTime(r["FFACTURA"]);

                    if (r["FFACTDES"] != System.DBNull.Value)
                        c.periodo_factura_desde = Convert.ToDateTime(r["FFACTDES"]);

                    if (r["FFACTHAS"] != System.DBNull.Value)
                        c.periodo_factura_hasta = Convert.ToDateTime(r["FFACTHAS"]);

                    if (r["TFACTURA"] != System.DBNull.Value)
                        c.tipo_factura = r["TFACTURA"].ToString();

                    if (r["AAJJ"] != System.DBNull.Value)
                        c.aajj = r["AAJJ"].ToString();

                    if (r["AGRECOBRO"] != System.DBNull.Value)
                        c.agrecobro = r["AGRECOBRO"].ToString();

                    if (r["PC"] != System.DBNull.Value)
                        c.pc = r["PC"].ToString();

                    if (r["FC"] != System.DBNull.Value)
                        c.fc = r["FC"].ToString();

                    List<EndesaEntity.cobros.Deuda_Tabla> o;
                    if (!dic.TryGetValue(c.nif, out o))
                    {
                        o = new List<EndesaEntity.cobros.Deuda_Tabla>();
                        o.Add(c);
                        dic.Add(c.nif, o);
                    }
                    else
                        o.Add(c);


                }
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "Deuda - Carga",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
            }

        }

        private List<EndesaEntity.cobros.Deuda_Tabla> GetDeuda(string nif)
        {
            List<EndesaEntity.cobros.Deuda_Tabla> o;
            if (dic.TryGetValue(nif, out o))
                return o;
            else
                return null;
        }

        public bool ExisteDeuda(string nif)
        {
            bool existe = false;
            List<EndesaEntity.cobros.Deuda_Tabla> o;
            if (dic.TryGetValue(nif, out o))
            {
                existe = true;
                lista = this.GetDeuda(nif);
            }

            return existe;
        }
    }
}
