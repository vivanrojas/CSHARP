using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class MotivosRechazos : EndesaEntity.contratacion.MotivosRechazo
    {

        public Dictionary<long, EndesaEntity.contratacion.MotivosRechazo> dic { get; set; }
        public MotivosRechazos()
        {
            dic = Carga();
        }


        private Dictionary<long, EndesaEntity.contratacion.MotivosRechazo> Carga()
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string f = "";

            Dictionary<long, EndesaEntity.contratacion.MotivosRechazo> d 
                = new Dictionary<long, EndesaEntity.contratacion.MotivosRechazo>();

            try
            {
                strSql = "SELECT r.CUPS, r.`Cliente Actualizado` AS Cliente_Actualizado,"
                    + " r.NumSolATR, r.`Fecha rechazo solicitud` AS Fecha_Rechazo_Solicitud,"
                    + " r.Tipo_solicitud, r.`Rechazo Pdte` AS Rechazo_Pdte, r.Motivo, r.Comentario"
                    + " FROM cont.rechazosatr r WHERE"
                    + " r.`Fecha rechazo solicitud` >= 20170301 AND"
                    + " (r.Motivo IS NULL OR r.Motivo = 'BAJA')"
                    + " ORDER BY r.`Fecha rechazo solicitud` DESC; ";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.MotivosRechazo c = new EndesaEntity.contratacion.MotivosRechazo();
                    if (r["CUPS"] != System.DBNull.Value)
                        c.cups = r["CUPS"].ToString();
                    if (r["Cliente_Actualizado"] != System.DBNull.Value)
                        c.clienteActualizado = r["Cliente_Actualizado"].ToString();
                    if (r["NumSolATR"] != System.DBNull.Value)
                        c.numSolAtr = Convert.ToInt64(r["NumSolATR"]);
                    if (r["Fecha_Rechazo_Solicitud"] != System.DBNull.Value)
                    {
                        f = r["Fecha_Rechazo_Solicitud"].ToString();
                        c.fechaRechazoSol =
                            new DateTime(Convert.ToInt32(f.Substring(0, 4)),
                            Convert.ToInt32(f.Substring(4, 2)),
                            Convert.ToInt32(f.Substring(6, 2)));
                    }
                    if (r["Tipo_solicitud"] != System.DBNull.Value)
                        c.tipoSolicitud = r["Tipo_solicitud"].ToString();
                    if (r["Rechazo_Pdte"] != System.DBNull.Value)
                        c.rechazoPdte = r["Rechazo_Pdte"].ToString();
                    if (r["Motivo"] != System.DBNull.Value)
                        c.motivos = r["Motivo"].ToString();
                    if (r["Comentario"] != System.DBNull.Value)
                        c.comentario = r["Comentario"].ToString();

                    d.Add(c.numSolAtr, c);

                }
                db.CloseConnection();

                return d;
            }catch(Exception e)
            {
                return null;
            }

        }

        public void Save()
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;

            try
            {
                

                strSql = "update rechazosatr set "
                    + " numSolAtr = " + this.numSolAtr;

                if (this.motivos != null && this.motivos != "")
                {
                    strSql += " ,Motivo = '" + this.motivos + "'";
                    ActualizaMotivo(this.numSolAtr, this.motivos);
                }
                else
                {
                    strSql += " ,Motivo = null";
                    ActualizaMotivo(this.numSolAtr, "");
                }
                    
                if (this.comentario != null && this.comentario != "")
                {
                    strSql += " ,Comentario = '" + this.comentario + "'";
                    ActualizaComentario(this.numSolAtr, this.comentario);
                }
                    

                strSql += " where NumSolATR = " + this.numSolAtr;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                command.ExecuteNonQuery();
                db.CloseConnection();
            }
            catch(Exception e)
            {

            }
        }

        private void ActualizaComentario(long numSolAtr, string comentario)
        {
            EndesaEntity.contratacion.MotivosRechazo o;
            if (dic.TryGetValue(numSolAtr, out o))
                o.comentario = comentario;
        }

        private void ActualizaMotivo(long numSolAtr, string motivo)
        {
            EndesaEntity.contratacion.MotivosRechazo o;
            if (dic.TryGetValue(numSolAtr, out o))
                o.motivos = motivo;
        }

    }
}
