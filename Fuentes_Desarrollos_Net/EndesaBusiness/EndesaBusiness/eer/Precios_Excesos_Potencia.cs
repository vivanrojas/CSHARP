using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.eer
{
    class Precios_Excesos_Potencia
    {
        Dictionary<string, List<EndesaEntity.facturacion.Precio_Exceso_Potencia>> dic;
        

        public Precios_Excesos_Potencia(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);            
        }

        private Dictionary<string, List<EndesaEntity.facturacion.Precio_Exceso_Potencia>> Carga(DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Dictionary<string, List<EndesaEntity.facturacion.Precio_Exceso_Potencia>> d =
                new Dictionary<string, List<EndesaEntity.facturacion.Precio_Exceso_Potencia>>();

            try
            {
                strSql = "select tarifa, tipo_equipo_medida, precio"
                    + " from eer_p_precio_excesos_potencia where"                    
                    + " (fecha_desde <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_hasta >= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Precio_Exceso_Potencia c = new EndesaEntity.facturacion.Precio_Exceso_Potencia();
                    if (r["tarifa"] != System.DBNull.Value)
                        c.tarifa = r["tarifa"].ToString();
                    if (r["tipo_equipo_medida"] != System.DBNull.Value)
                        c.tipo_equipo_medida = Convert.ToInt32(r["tipo_equipo_medida"]);
                    if (r["precio"] != System.DBNull.Value)
                        c.precio = Convert.ToDouble(r["precio"]);

                    List<EndesaEntity.facturacion.Precio_Exceso_Potencia> o;
                    if (!d.TryGetValue(c.tarifa, out o))
                    {
                        o = new List<EndesaEntity.facturacion.Precio_Exceso_Potencia>();
                        o.Add(c);
                        d.Add(c.tarifa, o);
                    }
                    else
                        o.Add(c);


                }
                db.CloseConnection();

                return d;
            }catch(Exception e)
            {

                MessageBox.Show("Error en Carga Precios Excesos Potencia: " + e.Message,               
                "Precios Excesos Potencia",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return null;
            }
        }

        public double GetValorPrecioExcesosPotencia(string tarifa, int tipo_punto_medida)
        {
            double resultado = 0;
            List<EndesaEntity.facturacion.Precio_Exceso_Potencia> o;
            if (dic.TryGetValue(tarifa, out o))
            {
                for (int i = 0; i < o.Count; i++)
                    if (o[i].tipo_equipo_medida == tipo_punto_medida)
                    {
                        resultado = o[i].precio;
                        break;
                    }

            }

            return resultado;

        }

    }
}
