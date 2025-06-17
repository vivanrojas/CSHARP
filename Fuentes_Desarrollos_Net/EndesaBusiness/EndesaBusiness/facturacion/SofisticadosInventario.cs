using EndesaBusiness.servidores;
using EndesaEntity.cnmc.V21_2019_12_17;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class SofisticadosInventario
    {
        public List<EndesaEntity.medida.PuntoSuministro> lista_pm { get; set; }
        Dictionary<string, EndesaEntity.medida.PuntoSuministro> dic { get; set; }
        public SofisticadosInventario(DateTime fd, DateTime fh, string cliente)
        {
            lista_pm = new List<EndesaEntity.medida.PuntoSuministro>();
            Carga(fd, fh, cliente);
        }

        public SofisticadosInventario(DateTime fd, DateTime fh)
        {
           dic = Carga(fd, fh);
        }

        private void Carga(DateTime fd, DateTime fh, string cliente)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            try
            {
                strSql = "SELECT i.cups20 FROM ag_inventario i"
                    + " WHERE i.cliente = '" + cliente.Trim() + "' and"
                    + " (i.fd <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " i.fh >= '" + fh.ToString("yyyy-MM-dd") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();
                    c.cups20 = r["cups20"].ToString();
                    lista_pm.Add(c);
                }
                db.CloseConnection();
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        private Dictionary<string, EndesaEntity.medida.PuntoSuministro> Carga(DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.medida.PuntoSuministro> d =
                new Dictionary<string, EndesaEntity.medida.PuntoSuministro>();

            try
            {
                strSql = "SELECT i.cups20 FROM ag_inventario i"
                    + " WHERE (i.fd <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " i.fh >= '" + fh.ToString("yyyy-MM-dd") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.PuntoSuministro c = new EndesaEntity.medida.PuntoSuministro();
                    c.cups20 = r["cups20"].ToString();

                    EndesaEntity.medida.PuntoSuministro o;
                    if (!d.TryGetValue(c.cups20, out o))
                        d.Add(c.cups20, c);

                }
                db.CloseConnection();
                return d;
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }

        }

        public bool Agora(string cups20)
        {
            EndesaEntity.medida.PuntoSuministro o;
            return dic.TryGetValue(cups20, out o);
        }

    }
}
