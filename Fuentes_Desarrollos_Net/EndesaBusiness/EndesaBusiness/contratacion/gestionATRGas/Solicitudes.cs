using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.gestionATRGas
{
    public class Solicitudes :  EndesaEntity.contratacion.gas.SolicitudesGas
    {
        public List<EndesaEntity.contratacion.gas.SolicitudesGas> lista { get; set; }        
        public Solicitudes(sigame.SIGAME sigame)
        {
            lista = new List<EndesaEntity.contratacion.gas.SolicitudesGas>();
            GetAll(sigame);
        }

        private void GetAll(sigame.SIGAME sigame)
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            

            try
            {
                strSql = "SELECT sd.id, sd.linea,"
                + " s.nif as cnifdnic, if(s.razon_social IS NULL, c.dapersoc, s.razon_social) as dapersoc, "
                + " c.distribuidora,"
                + " s.mail_remitente, s.fecha_mail,"
                + " s.cups, sd.producto, sd.qd, sd.fecha_inicio, sd.fecha_fin,"
                + " sd.id, sd.linea, sd.tarifa, sd.qi, sd.hora_inicio"
                + " FROM atrgas_solicitudes s"
                + " INNER JOIN atrgas_solicitudes_detalle sd ON"
                + " sd.id = s.id"
                + " LEFT OUTER JOIN cont.atrgas_contratos c on"
                + " c.cups20 = s.cups and"
                + " c.cnifdnic = s.nif"
                + " WHERE sd.aceptado = 'N'";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.gas.SolicitudesGas c = new EndesaEntity.contratacion.gas.SolicitudesGas();

                    c.cups = r["cups"].ToString();

                    if (r["id"] != System.DBNull.Value)
                        c.id = Convert.ToInt64(r["id"]);

                    if (r["linea"] != System.DBNull.Value)
                        c.linea = Convert.ToInt32(r["linea"]);

                    if (r["cnifdnic"] != System.DBNull.Value)
                        c.cnifdnic = r["cnifdnic"].ToString();

                    if (r["dapersoc"] != System.DBNull.Value)
                        c.dapersoc = r["dapersoc"].ToString();

                    if (r["distribuidora"] != System.DBNull.Value)
                        c.distribuidora = r["distribuidora"].ToString();
                    else
                        c.distribuidora = sigame.Distribuidora(c.cups);

                    if (r["mail_remitente"] != System.DBNull.Value)
                        c.remitente = r["mail_remitente"].ToString();

                    if (r["fecha_mail"] != System.DBNull.Value)
                        c.fecha_mail = Convert.ToDateTime(r["fecha_mail"]);

                    if (r["producto"] != System.DBNull.Value)
                        c.producto = r["producto"].ToString();

                    if (r["qd"] != System.DBNull.Value)
                        c.qd = Convert.ToDouble(r["qd"]);

                    if (r["fecha_inicio"] != System.DBNull.Value)
                        c.fecha_inicio = Convert.ToDateTime(r["fecha_inicio"]);

                    if (r["fecha_fin"] != System.DBNull.Value)
                        c.fecha_fin = Convert.ToDateTime(r["fecha_fin"]);

                    if (r["qi"] != System.DBNull.Value)
                        c.qi = Convert.ToDouble(r["qi"]);

                    if (r["hora_inicio"] != System.DBNull.Value)
                        c.hora_inicio = Convert.ToDateTime(r["hora_inicio"]);

                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();


                    lista.Add(c);
                }

                db.CloseConnection();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Solicitudes - GetAll",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }

        }

        public void ValidaSolicitud(int id, int linea)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "update atrgas_solicitudes_detalle set aceptado = 'S'"
                    + " where id = " + id + " and linea = " + linea;
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error a la hora de validar la solicitud",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }


        }

        public void CancelaSolicitud(int id, int linea)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                strSql = "update atrgas_solicitudes_detalle set aceptado = 'C'"
                    + " where id = " + id + " and linea = " + linea;
                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
               "Error a la hora de cancelar la solicitud",
               MessageBoxButtons.OK,
               MessageBoxIcon.Error);
            }


        }
    }
}
