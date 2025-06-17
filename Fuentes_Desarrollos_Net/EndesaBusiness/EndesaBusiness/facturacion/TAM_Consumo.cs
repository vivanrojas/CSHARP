using MySql.Data.MySqlClient;
using EndesaBusiness.servidores;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class TAM_Consumo
    {

        logs.Log ficheroLog;
        Dictionary<string, EndesaEntity.facturacion.TAM_Consumo> dic;
        public TAM_Consumo()
        {
            ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "FAC_TAM_Consumo");
            dic = new Dictionary<string, EndesaEntity.facturacion.TAM_Consumo>();
        }

        public void Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";


            try
            {
                strSql = "select cupsree, media_consumo"
                    + " from fo_tam_consumos";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.TAM_Consumo c = new EndesaEntity.facturacion.TAM_Consumo();
                    c.cupsree = r["cupsree"].ToString().Substring(0,20);
                    c.consumo = Convert.ToDouble(r["media_consumo"]);

                    EndesaEntity.facturacion.TAM_Consumo o;
                    if (!dic.TryGetValue(c.cupsree, out o))
                        dic.Add(c.cupsree, c);
                }
                db.CloseConnection();
            }
            catch(MySqlException e)
            {
                Console.WriteLine(e.Message);

            }
        }


        public void Proceso()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();
            int num_reg = 0;
            int total_reg = 0;
            double total_factura = 0;
            double total_consumo = 0;
            DateTime min_ffactdes = new DateTime(4999, 12, 31);
            DateTime max_ffacthas = DateTime.MinValue;
            int dias = 0;
            int dias_en_mes = 0;

            Dictionary<string, List<EndesaEntity.facturacion.TAM_Consumo>> dic
                = new Dictionary<string, List<EndesaEntity.facturacion.TAM_Consumo>>();

            try
            {
                //strSql = "SELECT t.CUPSREE,"
                //    + " MIN(t.FFACTDES) AS FFACTDES, MAX(t.FFACTHAS) AS FFACTHAS,"
                //    + " SUM(t.IFACTURA) AS IFACTURA, SUM(t.VCUOVAFA) AS VCUOVAFA"
                //    + " FROM fo_tam_consumos_tmp t"
                //    + " INNER JOIN (SELECT CUPSREE, MAX(FFACTDES) FFACTDES FROM fo_tam_consumos_tmp"
                //    + " GROUP BY CUPSREE) tt on"
                //    + "     tt.CUPSREE = t.CUPSREE"
                //    + " WHERE t.FFACTDES > tt.FFACTDES - INTERVAL 12 MONTH"
                //    + " AND t.CUPSREE like 'ES0021000000188551BV%'"
                //    + " GROUP BY t.CUPSREE, YEAR(t.FFACTDES), MONTH(t.FFACTDES)";


                strSql = "SELECT t.CUPSREE, t.FFACTDES, t.FFACTHAS, t.IFACTURA, t.VCUOVAFA"
                    + " FROM fo_tam_consumos_tmp t"
                    // + " where t.CUPSREE = 'ES0021000000188551BV0F'"
                    + " ORDER BY t.CUPSREE, t.FFACTDES DESC;";



                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.facturacion.TAM_Consumo c = new EndesaEntity.facturacion.TAM_Consumo();
                    c.cupsree = r["CUPSREE"].ToString();
                    c.ffactdes = Convert.ToDateTime(r["FFACTDES"]);
                    c.ffacthas = Convert.ToDateTime(r["FFACTHAS"]);
                    c.importe_factura = Convert.ToDouble(r["IFACTURA"]);
                    c.consumo = Convert.ToDouble(r["VCUOVAFA"]);


                    List<EndesaEntity.facturacion.TAM_Consumo> o;
                    if (!dic.TryGetValue(c.cupsree, out o))
                    {
                        o = new List<EndesaEntity.facturacion.TAM_Consumo>();
                        o.Add(c);
                        dic.Add(c.cupsree, o);
                    }
                    else
                    {
                        bool encontrado = false;
                        foreach(EndesaEntity.facturacion.TAM_Consumo p in o)
                        {
                            if(p.ffactdes.Year == c.ffactdes.Year && 
                                p.ffactdes.Month == c.ffactdes.Month)
                            {
                                encontrado = true;
                                p.ffactdes = c.ffactdes;
                                p.importe_factura = p.importe_factura + c.importe_factura;
                                p.consumo = p.consumo + c.consumo;
                            }
                        }

                        if(!encontrado)
                            o.Add(c);
                    }
                        

                }
                db.CloseConnection();


                foreach(KeyValuePair<string, List<EndesaEntity.facturacion.TAM_Consumo>> p in dic)
                {

                    

                    if (firstOnly)
                    {
                        sb.Append("replace into fo_tam_consumos (cupsree, media_importe_factura,"); 
                        sb.Append("media_consumo, ffactdes, ffacthas, num_meses_calculo) values ");
                        firstOnly = false;
                    }

                    num_reg++;
                    total_reg++;

                    total_factura = 0;
                    total_consumo = 0;

                    for(int i = 0; i < p.Value.Count; i++)
                    {
                        //dias_en_mes = DateTime.DaysInMonth(p.Value[i].ffactdes.Year, p.Value[i].ffactdes.Month);
                        //dias = Convert.ToInt32((p.Value[i].ffacthas - p.Value[i].ffactdes).TotalDays + 1);
                        //total_factura = total_factura + ((p.Value[i].importe_factura / dias) * dias_en_mes);
                        //total_consumo = total_consumo + ((p.Value[i].consumo / dias) * dias_en_mes);

                        total_factura = total_factura + p.Value[i].importe_factura;
                        total_consumo = total_consumo + p.Value[i].consumo;

                        min_ffactdes = (p.Value[i].ffactdes < min_ffactdes ? p.Value[i].ffactdes : min_ffactdes);
                        max_ffacthas = (p.Value[i].ffacthas > max_ffacthas ? p.Value[i].ffacthas : max_ffacthas);
                    }

                    sb.Append("('").Append(p.Value[0].cupsree).Append("',");

                    if(total_factura > 0)
                    {
                        double importe_total = Math.Round((total_factura / p.Value.Count), 2);
                        sb.Append(importe_total.ToString().Replace(",", ".")).Append(",");
                    }
                    else
                        sb.Append(0).Append(",");

                    if (total_consumo > 0)
                    {
                        double consumo_total = Math.Round((total_consumo / p.Value.Count), 2);
                        sb.Append(consumo_total.ToString().Replace(",", ".")).Append(",");
                    }
                    else
                        sb.Append(0).Append(",");


                    sb.Append("'").Append(min_ffactdes.ToString("yyyy-MM-dd")).Append("',");
                    sb.Append("'").Append(max_ffacthas.ToString("yyyy-MM-dd")).Append("',");

                    sb.Append(p.Value.Count).Append("),");


                    if (num_reg > 250)
                    {
                        Console.CursorLeft = 0;
                        Console.Write(total_reg.ToString("N0") + "/" + dic.Count.ToString("N0"));
                        db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        num_reg = 0;
                        strSql = "";
                        firstOnly = true;
                    }
                }

                if (num_reg > 0)
                {
                    Console.CursorLeft = 0;
                    Console.Write(total_reg.ToString("N0") + "/" + dic.Count.ToString("N0"));
                    db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    num_reg = 0;
                    strSql = "";
                    firstOnly = true;
                }


            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                ficheroLog.addError(ex.Message);
            }

        }

        public double GetConsumo(string cups20)
        {
            EndesaEntity.facturacion.TAM_Consumo o;
            if(dic.TryGetValue(cups20, out o))
                return o.consumo;
            else
                return 0;
        }

    }
}
