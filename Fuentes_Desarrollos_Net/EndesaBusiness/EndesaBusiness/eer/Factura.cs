using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.eer
{
    public class Factura
    {

        Dictionary<string, EndesaEntity.eer.Factura> dic_factura;
        public Factura()
        {
            dic_factura = Carga_UltimaFactura();
        }

        private Dictionary<string, EndesaEntity.eer.Factura> Carga_UltimaFactura()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.eer.Factura> d =
                new Dictionary<string, EndesaEntity.eer.Factura>();
            try
            {
                strSql = "SELECT f.nif, f.razon_social, f.direccion_suministro, f.cups20, f.tarifa"
                    + " FROM eer_facturas f ORDER BY f.fecha_consumo_desde DESC";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.eer.Factura c = new EndesaEntity.eer.Factura();
                    c.nif = r["nif"].ToString();
                    c.nombre_cliente = r["razon_social"].ToString();
                    c.direccion_suministro = r["direccion_suministro"].ToString();
                    c.tarifa = r["tarifa"].ToString();
                    c.cupsree = r["cups20"].ToString();

                    EndesaEntity.eer.Factura o;
                    if (!d.TryGetValue(c.cupsree, out o))
                    {
                        d.Add(c.cupsree, c);
                    }
                }
                db.CloseConnection();

                return d;

            }catch(Exception ex)
            {
                return null;
            }
        }


        public EndesaEntity.eer.Factura GetDatosCliente(string cups20)
        {
            EndesaEntity.eer.Factura o;
            if (dic_factura.TryGetValue(cups20, out o))
                return o;
            else
                return null;
        }

    }
}
