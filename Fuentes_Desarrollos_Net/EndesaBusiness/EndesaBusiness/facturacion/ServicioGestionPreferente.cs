using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.facturacion
{    
    public class ServicioGestionPreferente
    {
        Dictionary<string, double> dic;

        public string concepto { get; set; }
        public string calculo { get; set; }
        public double total { get; set; }

        public bool encontrado { get; set; }

        public ServicioGestionPreferente(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }

        private Dictionary<string, double> Carga(DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string cups20 = "";
            double precio = 0;

            Dictionary<string, double> d = new Dictionary<string, double>();

            try
            {
                strSql = "SELECT cups20, precio"
                    + " FROM cont.eer_p_servicio_gestion_preferente"
                    + " where (fecha_desde <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_hasta >= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {                    
                    cups20 = r["cups20"].ToString();
                    precio = Convert.ToDouble(r["precio"]);

                    d.Add(cups20, precio);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
              "ServicioGestionPreferente.Carga",
              MessageBoxButtons.OK,
              MessageBoxIcon.Error);
                return null;
            }
        }

        public void GetPrecio(string cups20, double energia)
        {
            double o;
            if(dic.TryGetValue(cups20, out o))
            {
                encontrado = true;
                concepto = "Servicio Gestión Prefer.Precio";
                calculo = (Math.Round(energia / 1000, 2)).ToString() + " MWh x "
                    + Math.Round(o, 2).ToString() + " Eur/MWh";
                total = Math.Round((energia / 1000) * o, 2);
            }
            else
            {
                encontrado = false;
            }
        }

    }
}
