using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.eer
{
    class RegularizacionPrecios
    {
        public Dictionary<string, List<EndesaEntity.facturacion.Regularizacion>> dic { get; set; }

        public RegularizacionPrecios(int mes_consumo)
        {
            dic = Carga(mes_consumo);
        }

        private Dictionary<string, List<EndesaEntity.facturacion.Regularizacion>> Carga(int mes_consumo)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
                      

            Dictionary<string, List<EndesaEntity.facturacion.Regularizacion>> d =
                new Dictionary<string, List<EndesaEntity.facturacion.Regularizacion>>();
                

            try
            {
                strSql = "SELECT cups20, texto_factura, importe FROM eer_regularizacion_precios"
                    + " where mes_consumo = " + mes_consumo;
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Regularizacion c = new EndesaEntity.facturacion.Regularizacion();

                    c.cups20 = r["cups20"].ToString();
                    c.texto_factura = r["texto_factura"].ToString();
                    c.importe = Convert.ToDouble(r["importe"]);

                    List<EndesaEntity.facturacion.Regularizacion> o;
                    if (!d.TryGetValue(c.cups20, out o))
                    {
                        o = new List<EndesaEntity.facturacion.Regularizacion>();
                        o.Add(c);
                        d.Add(c.cups20, o);
                    }
                    else
                        o.Add(c);

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
    }
}
