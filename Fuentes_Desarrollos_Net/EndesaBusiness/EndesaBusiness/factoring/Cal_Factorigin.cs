using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.factoring
{
    public  class Cal_Factorigin : EndesaEntity.factoring.CalendarioFactoring
    {
        public Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> dic;

        public Cal_Factorigin(string factoring)
        {
            dic = Carga(factoring);
            this.factoring = factoring;
        }

        public Cal_Factorigin()
        {
            dic = Carga();
            this.factoring = factoring;
        }

        private Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            try
            {
                Console.WriteLine("Cargando Calendario Factoring");
                Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> dic = 
                    new Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>>();
                strSql = "select factoring, facturas_desde, facturas_hasta, consumos_desde, consumos_hasta,"
                    + " fecha_ejecucion_desde, fecha_ejecucion_hasta, ejecutado,"
                    + " importe_min_factura, importe_min_factura_agrupada, bloque"
                    + " from ff_cal_informes_previos where"
                    + " (fecha_ejecucion_desde <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' and"
                    + " fecha_ejecucion_hasta >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "')"
                    + " order by bloque";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.factoring.CalendarioFactoring c = new EndesaEntity.factoring.CalendarioFactoring();
                    c.bloque = Convert.ToInt32(r["bloque"]);
                    c.factoring = r["factoring"].ToString();
                    c.facturas_desde = Convert.ToDateTime(r["facturas_desde"]);
                    c.facturas_hasta = Convert.ToDateTime(r["facturas_hasta"]);
                    c.consumos_desde = Convert.ToDateTime(r["consumos_desde"]);
                    c.consumos_hasta = Convert.ToDateTime(r["consumos_hasta"]);
                    c.fecha_ejecucion_desde = Convert.ToDateTime(r["fecha_ejecucion_desde"]);
                    c.fecha_ejecucion_hasta = Convert.ToDateTime(r["fecha_ejecucion_hasta"]);
                    c.ejecutado = r["ejecutado"].ToString() == "S" ? true : false;
                    c.importe_min_factura = Convert.ToDouble(r["importe_min_factura"]);
                    c.importe_min_factura_agrupada = Convert.ToDouble(r["importe_min_factura_agrupada"]);

                    List<EndesaEntity.factoring.CalendarioFactoring> o;
                    if (!dic.TryGetValue(c.factoring, out o))
                    {
                        List<EndesaEntity.factoring.CalendarioFactoring> lista = new List<EndesaEntity.factoring.CalendarioFactoring>();
                        lista.Add(c);
                        dic.Add(c.factoring, lista);
                    }
                    else
                        o.Add(c);



                }
                db.CloseConnection();
                return dic;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> Carga(string factoring)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            try
            {
                Console.WriteLine("Cargando Calendario Factoring");
                Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>> dic = 
                    new Dictionary<string, List<EndesaEntity.factoring.CalendarioFactoring>>();

                strSql = "select factoring, facturas_desde, facturas_hasta, consumos_desde, consumos_hasta," +
                    " fecha_ejecucion_desde, fecha_ejecucion_hasta, ejecutado,"
                    + " importe_min_factura, importe_min_factura_agrupada, bloque"
                    + " from ff_cal_informes_previos where"
                    + " factoring = " + factoring
                    // + " and facturas_hasta < '" + DateTime.Now.ToString("yyyy-MM-dd") + "'"
                    + " order by bloque";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.factoring.CalendarioFactoring c = new EndesaEntity.factoring.CalendarioFactoring();
                    c.bloque = Convert.ToInt32(r["bloque"]);
                    c.factoring = r["factoring"].ToString();
                    c.facturas_desde = Convert.ToDateTime(r["facturas_desde"]);
                    c.facturas_hasta = Convert.ToDateTime(r["facturas_hasta"]);
                    c.consumos_desde = Convert.ToDateTime(r["consumos_desde"]);
                    c.consumos_hasta = Convert.ToDateTime(r["consumos_hasta"]);
                    c.fecha_ejecucion_desde = Convert.ToDateTime(r["fecha_ejecucion_desde"]);
                    c.fecha_ejecucion_hasta = Convert.ToDateTime(r["fecha_ejecucion_hasta"]);
                    c.ejecutado = r["ejecutado"].ToString() == "S" ? true : false;
                    c.importe_min_factura = Convert.ToDouble(r["importe_min_factura"]);
                    c.importe_min_factura_agrupada = Convert.ToDouble(r["importe_min_factura_agrupada"]);

                    List<EndesaEntity.factoring.CalendarioFactoring> o;
                    if (!dic.TryGetValue(c.factoring, out o))
                    {
                        List<EndesaEntity.factoring.CalendarioFactoring> lista = new List<EndesaEntity.factoring.CalendarioFactoring>();
                        lista.Add(c);
                        dic.Add(c.factoring, lista);
                    }
                    else
                        o.Add(c);



                }
                db.CloseConnection();
                return dic;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public void MarcaEjecutado(string factoring, int bloque)
        {
            MySQLDB db;
            MySqlCommand command;
            string strSql = "";


            strSql = "update ff_cal_informes_previos set ejecutado = 'S'"
                + " where factoring = " + factoring
                + " and bloque = " + bloque;
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();
        }
    }
}
