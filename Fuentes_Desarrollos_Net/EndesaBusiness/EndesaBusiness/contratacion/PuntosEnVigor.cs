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
    public class PuntosEnVigor : EndesaEntity.contratacion.PS_AT_Tabla
    {
        public Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>> dic_cups13 { get; set; }
        public Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>> dic_cups20 { get; set; }

        private Dictionary<string, string> l_cups;
        public PuntosEnVigor()
        {
            dic_cups13 = new Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>>();
            dic_cups20 = new Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>>();
            l_cups = new Dictionary<string, string>();
            Carga(null);
        }

        public PuntosEnVigor(string empresa)
        {
            dic_cups13 = new Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>>();
            dic_cups20 = new Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>>();
            Carga(empresa);
        }

        private void Carga(string empresa)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            try
            {

                CargaListaCups();

                strSql = "select ps.EMPRESA, ps.IDU, ps.CUPS22, ps.NIF, ps.Cliente, TARIFA, provincia,"
                    + " ps.estadoCont as estadocontrato, ec.Descripcion estado_contrato"
                    + " from cont.PS_AT ps"
                    + " left outer join cont_estadoscontrato ec on"
                    + " ec.Cod_Estado = ps.estadoCont";

                if (empresa != null)
                    strSql += " where ps.EMPRESA = '" + empresa + "'";


                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.PS_AT_Tabla c = new EndesaEntity.contratacion.PS_AT_Tabla();
                    c.cups13 = r["IDU"].ToString();

                    if (r["EMPRESA"] != System.DBNull.Value)
                        c.empresa = r["EMPRESA"].ToString();
                    if (r["CUPS22"] != System.DBNull.Value)
                        c.cups22 = r["CUPS22"].ToString();
                    if (r["NIF"] != System.DBNull.Value)
                        c.cif = r["NIF"].ToString();
                    c.nombre_cliente = r["Cliente"].ToString();
                    if (r["TARIFA"] != System.DBNull.Value)
                        c.tarifa = r["TARIFA"].ToString();
                    if (r["estado_contrato"] != System.DBNull.Value)
                        c.estado_contrato_descripcion = r["estado_contrato"].ToString();
                    if (r["estadocontrato"] != System.DBNull.Value)
                        c.estado_contrato_id = Convert.ToInt32(r["estadocontrato"]);

                    if (c.cups22 == null || c.cups22 == "")
                        c.cups22 = GetInfoCups(c.cups13);

                    List<EndesaEntity.contratacion.PS_AT_Tabla> o;
                    if (!dic_cups13.TryGetValue(c.cups13, out o))
                    {
                        List<EndesaEntity.contratacion.PS_AT_Tabla> lista = new List<EndesaEntity.contratacion.PS_AT_Tabla>();
                        lista.Add(c);
                        dic_cups13.Add(c.cups13, lista);
                    }
                    else
                    {
                        o.Add(c);
                    }


                    List<EndesaEntity.contratacion.PS_AT_Tabla> g;

                    if (c.cups22 != "")
                    {
                        if (!dic_cups20.TryGetValue(c.cups22.Substring(0, 20), out g))
                        {
                            List<EndesaEntity.contratacion.PS_AT_Tabla> lista = new List<EndesaEntity.contratacion.PS_AT_Tabla>();
                            lista.Add(c);

                            dic_cups20.Add(c.cups22.Substring(0, 20), lista);
                        }
                        else
                        {
                            g.Add(c);
                        }
                    }


                }

                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                   "PuntosEnVigor_PS_AT - Carga",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error);

            }
        }


        private void CargaListaCups()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            try
            {

                strSql = "select substr(cups15,1,13) as cups13, substr(cups22,1,20) as cups20 from med.dt_cups group by substr(cups15,1,13)";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    string a;
                    if (r["cups13"] != System.DBNull.Value && r["cups20"] != System.DBNull.Value)
                        if (!l_cups.TryGetValue(r["cups13"].ToString(), out a))
                            l_cups.Add(r["cups13"].ToString(), r["cups20"].ToString());
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                    "CargaListaCups",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
            }
        }

        private string GetInfoCups(string cups13)
        {
            string a;
            if (l_cups.TryGetValue(cups13, out a))
                return a;
            else
                return "";

        }

        public bool Existe(string cups13o20)
        {
            bool existe = false;
            List<EndesaEntity.contratacion.PS_AT_Tabla> o;
            if (cups13o20.Length == 13)
            {
                existe = dic_cups13.TryGetValue(cups13o20, out o);

            }

            else
                existe = dic_cups20.TryGetValue(cups13o20, out o);


            return existe;
        }
    }
}
