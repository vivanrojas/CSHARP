using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion
{
    public class TiposMotivosRechazo
    {
        public List<EndesaEntity.contratacion.TiposMotivoRechazo> list { get; set; }

        public TiposMotivosRechazo()
        {
            list = Carga();
        }

        private List<EndesaEntity.contratacion.TiposMotivoRechazo> Carga() 
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            List<EndesaEntity.contratacion.TiposMotivoRechazo> l
                = new List<EndesaEntity.contratacion.TiposMotivoRechazo>();

            try
            {
                strSql = "SELECT motivo, solucion FROM param_motivos_rechazo";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.contratacion.TiposMotivoRechazo c = new EndesaEntity.contratacion.TiposMotivoRechazo();
                    if (r["motivo"] != System.DBNull.Value)
                        c.motivo = r["motivo"].ToString();
                    if (r["motivo"] != System.DBNull.Value)
                        c.solucion = r["solucion"].ToString();

                    l.Add(c);

                }
                db.CloseConnection();
                    return l;
            }catch(Exception e)
            {
                return null;
            }


        }





    }
}
