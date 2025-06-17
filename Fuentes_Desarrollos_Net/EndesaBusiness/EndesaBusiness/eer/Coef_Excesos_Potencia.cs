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
    public class Coef_Excesos_Potencia
    {
        Dictionary<string, List<EndesaEntity.facturacion.Coeficiente_Excesos_Potencia>> dic;

        

        public Coef_Excesos_Potencia(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }
                
        private Dictionary<string, List<EndesaEntity.facturacion.Coeficiente_Excesos_Potencia>> Carga(DateTime fd, DateTime fh)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            Dictionary<string, List<EndesaEntity.facturacion.Coeficiente_Excesos_Potencia>> d =
                new Dictionary<string, List<EndesaEntity.facturacion.Coeficiente_Excesos_Potencia>>();

            string tarifa = "";
            

            try
            {
                strSql = "SELECT tarifa, periodo_tarifario, fecha_desde, fecha_hasta, valor, f_ult_mod"
                    + " FROM cont.eer_p_coef_excesos_potencias where"                    
                    + " (fecha_desde <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_hasta >= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    if (r["tarifa"] != System.DBNull.Value)
                    {
                        tarifa = r["tarifa"].ToString();
                            
                        EndesaEntity.facturacion.Coeficiente_Excesos_Potencia c = new EndesaEntity.facturacion.Coeficiente_Excesos_Potencia();
                        c.periodo = Convert.ToInt32(r["periodo_tarifario"]);                        
                        if (r["valor"] != System.DBNull.Value)                        
                            c.valor = Convert.ToDouble(r["valor"]);

                        List<EndesaEntity.facturacion.Coeficiente_Excesos_Potencia> o;
                        if (!d.TryGetValue(tarifa, out o))
                        {
                            o = new List<EndesaEntity.facturacion.Coeficiente_Excesos_Potencia>();
                            o.Add(c);
                            d.Add(tarifa, o);
                        }
                        else
                            o.Add(c);

                    }

                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error en Carga Coeficiente Excesos Potencia: " + e.Message,
                "Coeficiente Excesos Potencia",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return null;
            }
        }
        public double GetValorExcesosPotencia(string tarifa, int periodoTarifario)
        {
            double resultado = 0;
            List<EndesaEntity.facturacion.Coeficiente_Excesos_Potencia> o;
            if (dic.TryGetValue(tarifa, out o))
            {
                for (int i = 0; i < o.Count; i++)
                    if (o[i].periodo == periodoTarifario)
                    {
                        resultado = o[i].valor;
                        break;
                    }
                        
            }

            return resultado;
            
        }


    }
}
