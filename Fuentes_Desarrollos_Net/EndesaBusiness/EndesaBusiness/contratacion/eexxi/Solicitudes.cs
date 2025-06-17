using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.contratacion.eexxi
{
    public class Solicitudes : EndesaEntity.contratacion.xxi.XML_Datos
    {

        public Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> dic;
        public Solicitudes()
        {

        }

        public Solicitudes(string proceso, string paso)
        {
            dic = CargaSolicitudes(proceso, paso);
        }


        private Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> CargaSolicitudes(string proceso, string paso)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud> d =
                new Dictionary<string, EndesaEntity.contratacion.xxi.Cups_Solicitud>();

            try
            {

                if(proceso == "T1" && paso == "01")
                {


                    strSql = "SELECT s.CUPS, s.CodigoDeSolicitud, s.FechaSolicitud,"
                   + " s.CodigoDelProceso, s.CodigoDePaso"
                   + " FROM eexxi_solicitudes_t101 s WHERE"
                   + " s.CodigoDelProceso = '" + proceso + "' AND"
                   + " s.CodigoDePaso = '" + paso + "'"
                   + " order by s.FechaSolicitud desc";
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        EndesaEntity.contratacion.xxi.Cups_Solicitud c = new EndesaEntity.contratacion.xxi.Cups_Solicitud();

                        c.proceso = r["CodigoDelProceso"].ToString();
                        c.paso = r["CodigoDePaso"].ToString();

                        if (r["CUPS"] != System.DBNull.Value)
                            c.cups = r["CUPS"].ToString();

                        if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                            c.solicitud = r["CodigoDeSolicitud"].ToString();

                        if (r["FechaSolicitud"] != System.DBNull.Value)
                            c.fecha_solicitud = Convert.ToDateTime(r["FechaSolicitud"]);

                        EndesaEntity.contratacion.xxi.Cups_Solicitud o;
                        if (!d.TryGetValue(c.cups, out o))
                        {
                            d.Add(c.cups, c);
                        }

                    }
                    db.CloseConnection();




                    



                    

                   

                }
                else
                {
                    strSql = "SELECT s.CUPS, s.CodigoDeSolicitud, s.FechaSolicitud,"
                                        + " s.CodigoDelProceso, s.CodigoDePaso"
                                         + " FROM eexxi_solicitudes_tmp s WHERE"
                                         + " s.CodigoDelProceso = '" + proceso + "' AND"
                                         + " s.CodigoDePaso = '" + paso + "'"
                                         + " order by s.FechaSolicitud desc";
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        EndesaEntity.contratacion.xxi.Cups_Solicitud c = new EndesaEntity.contratacion.xxi.Cups_Solicitud();

                        c.proceso = r["CodigoDelProceso"].ToString();
                        c.paso = r["CodigoDePaso"].ToString();

                        if (r["CUPS"] != System.DBNull.Value)
                            c.cups = r["CUPS"].ToString();

                        if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                            c.solicitud = r["CodigoDeSolicitud"].ToString();

                        if (r["FechaSolicitud"] != System.DBNull.Value)
                            c.fecha_solicitud = Convert.ToDateTime(r["FechaSolicitud"]);

                        EndesaEntity.contratacion.xxi.Cups_Solicitud o;
                        if (!d.TryGetValue(c.cups, out o))
                        {
                            d.Add(c.cups, c);
                        }

                    }
                    db.CloseConnection();

                    strSql = "SELECT s.CUPS, s.CodigoDeSolicitud, s.FechaSolicitud,"
                   + " s.CodigoDelProceso, s.CodigoDePaso"
                   + " FROM eexxi_solicitudes s WHERE"
                   + " s.CodigoDelProceso = '" + proceso + "' AND"
                   + " s.CodigoDePaso = '" + paso + "'"
                   + " order by s.FechaSolicitud desc";
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                    command = new MySqlCommand(strSql, db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        EndesaEntity.contratacion.xxi.Cups_Solicitud c = new EndesaEntity.contratacion.xxi.Cups_Solicitud();

                        c.proceso = r["CodigoDelProceso"].ToString();
                        c.paso = r["CodigoDePaso"].ToString();

                        if (r["CUPS"] != System.DBNull.Value)
                            c.cups = r["CUPS"].ToString();

                        if (r["CodigoDeSolicitud"] != System.DBNull.Value)
                            c.solicitud = r["CodigoDeSolicitud"].ToString();

                        if (r["FechaSolicitud"] != System.DBNull.Value)
                            c.fecha_solicitud = Convert.ToDateTime(r["FechaSolicitud"]);

                        EndesaEntity.contratacion.xxi.Cups_Solicitud o;
                        if (!d.TryGetValue(c.cups, out o))
                        {
                            d.Add(c.cups, c);
                        }

                    }
                    db.CloseConnection();
                }
                             

                return d;

            }
            catch(Exception ex)
            {
                return null;
            }
        }

        public void Update()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;


            try
            {
                strSql = "update eexxi_solicitudes_tmp set CUPS = '" + this.cups + "'";

                if (this.linea1DeLaDireccionExterna != null)
                    strSql += " ,Linea1DeLaDireccionExterna = '" + this.linea1DeLaDireccionExterna + "'";
                else
                    strSql += " ,Linea1DeLaDireccionExterna = null";

                if (this.linea2DeLaDireccionExterna != null)
                    strSql += " ,Linea2DeLaDireccionExterna = '" + this.linea2DeLaDireccionExterna + "'";
                else
                    strSql += " ,Linea2DeLaDireccionExterna = null";

                if (this.linea3DeLaDireccionExterna != null)
                    strSql += " ,Linea3DeLaDireccionExterna = '" + this.linea3DeLaDireccionExterna + "'";
                else
                    strSql += " ,Linea3DeLaDireccionExterna = null";

                if (this.linea4DeLaDireccionExterna != null)
                    strSql += " ,Linea4DeLaDireccionExterna = '" + this.linea4DeLaDireccionExterna + "'";
                else
                    strSql += " ,Linea4DeLaDireccionExterna = null";

                strSql += " where CodigoDeSolicitud = '" + this.codigoDeSolicitud + "'";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();

                strSql = "update eexxi_solicitudes set CUPS = '" + this.cups + "'";

                if (this.linea1DeLaDireccionExterna != null)
                    strSql += " ,Linea1DeLaDireccionExterna = '" + this.linea1DeLaDireccionExterna + "'";
                else
                    strSql += " ,Linea1DeLaDireccionExterna = null";

                if (this.linea2DeLaDireccionExterna != null)
                    strSql += " ,Linea2DeLaDireccionExterna = '" + this.linea2DeLaDireccionExterna + "'";
                else
                    strSql += " ,Linea2DeLaDireccionExterna = null";

                if (this.linea3DeLaDireccionExterna != null)
                    strSql += " ,Linea3DeLaDireccionExterna = '" + this.linea3DeLaDireccionExterna + "'";
                else
                    strSql += " ,Linea3DeLaDireccionExterna = null";

                if (this.linea4DeLaDireccionExterna != null)
                    strSql += " ,Linea4DeLaDireccionExterna = '" + this.linea4DeLaDireccionExterna + "'";
                else
                    strSql += " ,Linea4DeLaDireccionExterna = null";

                strSql += " where CodigoDeSolicitud = '" + this.codigoDeSolicitud + "'";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();




            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                        "Solicitudes.Update",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Error);
            }


        }

        public void GetSolicitud(string cups22)
        {
            EndesaEntity.contratacion.xxi.Cups_Solicitud o;
            if (dic.TryGetValue(cups22, out o))
            {
                this.codigoDeSolicitud = o.solicitud;
                this.existe = true;
            }
        }

    }
}
