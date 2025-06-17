using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.cuadro_mando
{
    class Scoring
    {
        Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Scoring> dic;
        public Scoring()
        {
            dic = Carga();
        }

        private Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Scoring> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            StringBuilder sb = new StringBuilder();
            

            Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Scoring> d
                = new Dictionary<string, EndesaEntity.facturacion.cuadroDeMando.Scoring>();

            try
            {
                sb.Append("SELECT c.cif, c.prioridad FROM covid19_clientes c");
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(sb.ToString(), db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.cuadroDeMando.Scoring c = 
                        new EndesaEntity.facturacion.cuadroDeMando.Scoring();
                    c.nif = r["cif"].ToString();
                    c.prioridad = r["prioridad"].ToString();

                    EndesaEntity.facturacion.cuadroDeMando.Scoring o;
                    if (!d.TryGetValue(c.nif, out o))
                        d.Add(c.nif, c);
                }

                db.CloseConnection();
                return d;
            }catch(Exception e)
            {
                return null;
            }
        }

        public string Prioridad(string nif)
        {
            EndesaEntity.facturacion.cuadroDeMando.Scoring o;
            if (dic.TryGetValue(nif, out o))
                return o.prioridad;
            else
                return null;
        }
    }
}
