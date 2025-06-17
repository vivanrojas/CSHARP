using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EndesaEntity;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class Distribuidoras : Table_atrgas_distribuidoras
    {
        logs.Log ficheroLog;
        public Dictionary<string, Table_atrgas_distribuidoras> l_distribuidoras { get; set; }
        EndesaBusiness.contratacion.gestionATRGas.ContratosGas contratosGas;

        public Distribuidoras(bool cargaInventario)
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_Distribuidoras");
            l_distribuidoras = new Dictionary<string, Table_atrgas_distribuidoras>();
            GetAll();
            contratosGas = new EndesaBusiness.contratacion.gestionATRGas.ContratosGas();
        }

        public Distribuidoras()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "CON_Distribuidoras");
            l_distribuidoras = new Dictionary<string, Table_atrgas_distribuidoras>();
            
        }

        public void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "update atrgas_distribuidoras set";

                strSql += " nombre = '" + this.nombre + "',"
                    + "fecha_desde = '" + this.fecha_desde.ToString("yyyy-MM-dd") + "',"
                    + "fecha_hasta = '" + this.fecha_hasta.ToString("yyyy-MM-dd") + "',"
                    + "email = '" + this.email + "',"
                    + "last_update_by = '" + System.Environment.UserName + "'"
                    + " where codigo = '" + this.codigo + "'";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en la actualización de datos",
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
                strSql = "insert into atrgas_distribuidoras (codigo, nombre, fecha_desde,"
                    + "fecha_hasta, email, created_by, created_date) values ";

                strSql += "('" + codigo + "',"
                    + "'" + nombre + "',"
                    + "'" + fecha_desde.ToString("yyyy-MM-dd") + "',"
                    + "'" + fecha_hasta.ToString("yyyy-MM-dd") + "',"
                    + "'" + email + "',"
                    + "'" + System.Environment.UserName + "',"
                    + "'" + DateTime.Now.ToString("yyyy-MM-dd") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                MessageBox.Show("La distribuidora " + this.codigo + " se ha añadido correctamente",
                  "Distribuidora añadida",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information);


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
                strSql = "delete from atrgas_distribuidoras where codigo = '" + codigo + "';";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error en el borrado de datos",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }
        }

        private void Read()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                strSql = "select codigo, nombre, codigo_xml, tramitacion, fecha_desde, fecha_hasta, email, created_by, created_date, last_update_by, last_update_date"
                    + " from atrgas_distribuidoras";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Table_atrgas_distribuidoras c = new Table_atrgas_distribuidoras();

                    if (r["codigo"] != System.DBNull.Value)
                        c.codigo = r["codido"].ToString();
                    if (r["tramitacion"] != System.DBNull.Value)
                        c.tramitacion = r["tramitacion"].ToString();
                    if (r["codigo_xml"] != System.DBNull.Value)
                        c.codigo_xml = r["codigo_xml"].ToString();
                    if (r["nombre"] != System.DBNull.Value)
                        c.nombre = r["nombre"].ToString();
                    if (r["fecha_desde"] != System.DBNull.Value)
                        c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);
                    if (r["fecha_hasta"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);
                    if (r["email"] != System.DBNull.Value)
                        c.email = r["email"].ToString();
                    if (r["created_by"] != System.DBNull.Value)
                        c.created_by = r["created_by"].ToString();
                    if (r["created_date"] != System.DBNull.Value)
                        c.creation_date = Convert.ToDateTime(r["created_date"]);
                    if (r["last_update_by"] != System.DBNull.Value)
                        c.last_update_by = r["last_update_by"].ToString();
                    if (r["last_update_date"] != System.DBNull.Value)
                        c.last_update_date = Convert.ToDateTime(r["last_update_date"]);



                    Table_atrgas_distribuidoras a;
                    if (!l_distribuidoras.TryGetValue(codigo, out a))
                        l_distribuidoras.Add(codigo, a);
                }
                db.CloseConnection();
            }
            catch (Exception e)
            {

            }
        }

        private void GetAll()
        {

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;            

            try
            {
                strSql = "select codigo, nombre, codigo_xml, tramitacion, fecha_desde, fecha_hasta, email,"
                + " created_by, created_date, last_update_by, last_update_date"
                + " from atrgas_distribuidoras";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    Table_atrgas_distribuidoras c = new Table_atrgas_distribuidoras();
                    c.codigo = r["codigo"].ToString().ToUpper();

                    if (r["nombre"] != System.DBNull.Value)
                        c.nombre = r["nombre"].ToString().ToUpper();

                    if (r["codigo_xml"] != System.DBNull.Value)
                        c.codigo_xml = r["codigo_xml"].ToString();

                    if (r["tramitacion"] != System.DBNull.Value)
                        c.tramitacion = r["tramitacion"].ToString();

                    if (r["fecha_desde"] != System.DBNull.Value)
                        c.fecha_desde = Convert.ToDateTime(r["fecha_desde"]);

                    if (r["fecha_hasta"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["fecha_hasta"]);

                    if (r["email"] != System.DBNull.Value)
                        c.email = r["email"].ToString();

                    if (r["created_by"] != System.DBNull.Value)
                        c.created_by = r["created_by"].ToString();

                    if (r["created_date"] != System.DBNull.Value)
                        c.creation_date = Convert.ToDateTime(r["created_date"]);

                    if (r["last_update_by"] != System.DBNull.Value)
                        c.last_update_by = r["last_update_by"].ToString();

                    c.last_update_date = Convert.ToDateTime(r["last_update_date"]);

                    l_distribuidoras.Add(c.codigo, c);
                }

                db.CloseConnection();
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Azucarera.DescargaArchivo: " + e.Message);
            }

        }

        public string Codigo_XML_CNMC_Distribuidora(string nombre_distribuidora)
        {
            Table_atrgas_distribuidoras o;
            if (l_distribuidoras.TryGetValue(nombre_distribuidora.ToUpper(), out o))
                return o.codigo_xml;
            else
                return "";
        }

        public string GetMailDistruidora(string distribuidora)
        {
            Table_atrgas_distribuidoras o;
            if (l_distribuidoras.TryGetValue(distribuidora, out o))
                return o.email;
            else
                return null;
        }

        public string GetTramitacion(string distribuidora, string nif, string cups20)
        {

            string tramitacion = "Distribuidora";

            if (contratosGas.ExisteContrato(nif, cups20))
            {
                contratosGas.GetRecord(nif, cups20);
                tramitacion = contratosGas.tramitacion;
            }


            if (tramitacion == "Distribuidora")
            {
                Table_atrgas_distribuidoras o;
                if (l_distribuidoras.TryGetValue(distribuidora.Trim().ToUpper(), out o))
                    tramitacion = o.tramitacion;
                else
                    return null;
            }


            return tramitacion;
        }


    }
}
