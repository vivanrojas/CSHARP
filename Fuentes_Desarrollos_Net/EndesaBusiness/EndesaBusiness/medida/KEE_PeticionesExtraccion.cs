using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class KEE_PeticionesExtraccion
    {
        List<EndesaEntity.medida.KEE_PeticionExtraccion> lista;
        public KEE_PeticionesExtraccion()
        {

        }

        private List<EndesaEntity.medida.KEE_PeticionExtraccion> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            List<EndesaEntity.medida.KEE_PeticionExtraccion> l
                = new List<EndesaEntity.medida.KEE_PeticionExtraccion>();

            try
            {
                strSql = "select cups, fecha_desde, fecha_hasta,"
                    + " origen, tipo_curva, motivo, usuario, fecha_envio"
                    + " from kee_peticiones_extraccion";
                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.KEE_PeticionExtraccion c = new EndesaEntity.medida.KEE_PeticionExtraccion();
                    if (r["cups"] != System.DBNull.Value)
                        c.cups = r["cups"].ToString();


                }
                db.CloseConnection();
                return lista;
            }
            catch(Exception e)
            {
                return null;
            }
        }
    }
}
