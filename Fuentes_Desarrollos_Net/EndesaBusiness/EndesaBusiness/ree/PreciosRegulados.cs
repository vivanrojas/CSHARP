using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.ree
{
    public class PreciosRegulados
    {
        Dictionary<string, EndesaEntity.ree.PreciosRegulados> dic;

        public PreciosRegulados()
        {

        }
        public PreciosRegulados(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }

        private Dictionary<string, EndesaEntity.ree.PreciosRegulados> Carga(DateTime fd, DateTime fh)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, EndesaEntity.ree.PreciosRegulados> d
                = new Dictionary<string, EndesaEntity.ree.PreciosRegulados>();

            try
            {
                strSql = "SELECT Termino, Tarifa, FechaDesde, FechaHasta, Unidad, PT1, PT2, PT3, PT4, PT5, PT6"
                    + " FROM cont.eer_componentes_regulados"
                    + " where (FechaDesde <= '" + fh.ToString("yyyy-MM-dd") + "' and"
                    + " FechaHasta >= '" + fd.ToString("yyyy-MM-dd") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    EndesaEntity.ree.PreciosRegulados c = new EndesaEntity.ree.PreciosRegulados();
                    if (r["Termino"] != System.DBNull.Value)
                        c.termino = r["Termino"].ToString();

                    if (r["Tarifa"] != System.DBNull.Value)
                        c.tarifa = r["Tarifa"].ToString();

                    if (r["FechaDesde"] != System.DBNull.Value)
                        c.fecha_desde = Convert.ToDateTime(r["FechaDesde"]);

                    if (r["FechaHasta"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["FechaHasta"]);

                    if (r["Unidad"] != System.DBNull.Value)
                        c.unidad = r["Unidad"].ToString();

                    for (int i = 1; i <= 6; i++)
                    {
                        if (r["PT" + i] != System.DBNull.Value)
                            c.periodo_tarifario[i] = Convert.ToDouble(r["PT" + i]);
                    }


                    EndesaEntity.ree.PreciosRegulados o;
                    if (!d.TryGetValue(c.termino + c.tarifa, out o))
                        d.Add(c.termino + c.tarifa, c);
                    


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

        public EndesaEntity.ree.PreciosRegulados GetPrecioRegulado(string termino, string tarifa)
        {
            EndesaEntity.ree.PreciosRegulados o;
            if (dic.TryGetValue(termino + tarifa, out o))
                return o;          
            else
                return null;
        }

        public EndesaEntity.ree.PreciosRegulados GetPrecioRegulado(string termino1, string termino2, string tarifa)
        {
            EndesaEntity.ree.PreciosRegulados r = new EndesaEntity.ree.PreciosRegulados();
            EndesaEntity.ree.PreciosRegulados o;
            
            if (dic.TryGetValue(termino1 + tarifa, out o))
            {
                r.fecha_desde = o.fecha_desde;
                r.fecha_hasta = o.fecha_hasta;
                r.tarifa = o.tarifa;
                r.unidad = o.unidad;
                r.termino = o.termino;

                for (int i = 1; i <= 6; i++)
                    r.periodo_tarifario[i] = o.periodo_tarifario[i];
            }

            if (dic.TryGetValue(termino2 + tarifa, out o))
            {
                                
                for (int i = 1; i <= 6; i++)
                    r.periodo_tarifario[i] = r.periodo_tarifario[i] + o.periodo_tarifario[i];
            }

            return r;

        }

        public string Texto_Componentes_Regulados(double importe_potencias, double importe_excesos_potencia,
            EndesaEntity.punto_suministro.Tarifa tarifa, double[] energia_activa, double importe_energia_reactiva)
        {
            string texto = "";
            double total = 0;
            double total_energia_activa = 0;

            EndesaEntity.ree.PreciosRegulados c = GetPrecioRegulado("TerminoEnergia", tarifa.tarifa);
            for (int i = 1; i <= tarifa.numPeriodosTarifarios; i++)
                total_energia_activa = total_energia_activa + (energia_activa[i] * c.periodo_tarifario[i]);

            total = total_energia_activa + importe_potencias + importe_excesos_potencia + importe_energia_reactiva;

            texto = "Incluido en el importe facturado está el coste del peaje de acceso que ha sido de "
                + string.Format("{0:#,##0.00}", total) + " € ("
                + string.Format("{0:#,##0.00}", importe_potencias + importe_excesos_potencia) + " € potencia, "
                + string.Format("{0:#,##0.00}", total_energia_activa) + " € por energía activa y "
                + string.Format("{0:#,##0.00}", importe_energia_reactiva) + " € por energía reactiva). "
                + "Precio del peaje de acceso publicados en la Orden TEC/1258/2019 (BOE 28-12-2019).";
            return texto;

        }

        public string Texto_Componentes_Regulados_2021(double TP, double TE,
            double cargos_potencia, double cargos_activa,
            EndesaEntity.punto_suministro.Tarifa tarifa, double[] energia_activa, double importe_energia_reactiva)
        {
            string texto = "";            
            double total_energia_activa = 0;                        

            texto = "Incluido en el importe facturado está el coste del peaje de transporte y distribución, que ha sido de "
                + string.Format("{0:#,##0.00}", TP + TE + importe_energia_reactiva) + " € ("
                + string.Format("{0:#,##0.00}", TP) + " € potencia, "
                + string.Format("{0:#,##0.00}", TE) + " € por energía activa "
                + string.Format("{0:#,##0.00}", importe_energia_reactiva) + " € por energía reactiva) y de los cargos, "
                + "que ha sido de " + string.Format("{0:#,##0.00}", cargos_potencia + cargos_activa) + " € ("
                + string.Format("{0:#,##0.00}", cargos_potencia) + " € potencia "
                + string.Format("{0:#,##0.00}", cargos_activa) + " € energía activa). "
                + "Los precios de peajes de transporte y distribución han sido publicados en la Resolución de 18 de marzo "
                + "de 2021, de la CNMC (BOE 23-03-2021) y los de cargos en la Orden TED/371/2021 (BOE 22-04-2021).";
            return texto;

        }

        public string Texto_Componentes_Regulados_2022(double TP, double TE,
            double cargos_potencia, double cargos_activa,
            EndesaEntity.punto_suministro.Tarifa tarifa, double[] energia_activa, double importe_energia_reactiva)
        {
            string texto = "";
            double total_energia_activa = 0;

            texto = "Incluido en el importe facturado está el coste del peaje de transporte y distribución, que ha sido de "
                + string.Format("{0:#,##0.00}", TP + TE + importe_energia_reactiva) + " € ("
                + string.Format("{0:#,##0.00}", TP) + " € potencia, "
                + string.Format("{0:#,##0.00}", TE) + " € por energía activa "
                + string.Format("{0:#,##0.00}", importe_energia_reactiva) + " € por energía reactiva) y de los cargos, "
                + "que ha sido de " + string.Format("{0:#,##0.00}", cargos_potencia + cargos_activa) + " € ("
                + string.Format("{0:#,##0.00}", cargos_potencia) + " € potencia "
                + string.Format("{0:#,##0.00}", cargos_activa) + " € energía activa). "
                + "Los precios de peajes de transporte y distribución han sido publicados en la Resolución de 16 de diciembre "
                + "de 2021, de la CNMC (BOE 22-12-2021) y los de cargos en la Orden TED/1484/2021 (BOE 28-12-2021).";
            return texto;

        }

        public string Texto_Componentes_Regulados_20220331(double TP, double TE,
            double cargos_potencia, double cargos_activa,
            EndesaEntity.punto_suministro.Tarifa tarifa, double[] energia_activa, double importe_energia_reactiva)
        {
            string texto = "";
            double total_energia_activa = 0;

            texto = "Incluido en el importe facturado está el coste del peaje de transporte y distribución, que ha sido de "
                + string.Format("{0:#,##0.00}", TP + TE + importe_energia_reactiva) + " € ("
                + string.Format("{0:#,##0.00}", TP) + " € potencia, "
                + string.Format("{0:#,##0.00}", TE) + " € por energía activa "
                + string.Format("{0:#,##0.00}", importe_energia_reactiva) + " € por energía reactiva) y de los cargos, "
                + "que ha sido de " + string.Format("{0:#,##0.00}", cargos_potencia + cargos_activa) + " € ("
                + string.Format("{0:#,##0.00}", cargos_potencia) + " € potencia "
                + string.Format("{0:#,##0.00}", cargos_activa) + " € energía activa). "
                + "Los precios de peajes de transporte y distribución han sido publicados en la Resolución de 30 de diciembre "
                + "de 2021, de la CNMC (BOE 22-12-2021) y los de cargos en la Orden TED/1484/2021 (BOE 28-12-2021),"
                + " actualizado por RDL 6-2022 (30/03/2022).";
            return texto;

        }

        public string Texto_Actual()
        {
            string texto = "";

            texto = "La estructura de su peaje pasará a ser la que le corresponda según "
                + "lo regulado en los Artículos 6, 7 y 9 de la Circular 3/2020 de la CNMC "
                + "publicada en el BOE del 24 de enero de 2020, en el plazo y en las condiciones "
                + "establecidas en dicha Circular y en la legislación vigente.";
                return texto;
        }

        public string Texto_RDL17_2021(double descuento)
        {
            string texto = "";

            texto = "El importe facturado incluye un descuento de "
                + string.Format("{0:#,##0.00}", descuento) + " euros "
                + "derivado de los nuevos precios de los cargos publicados en el RDL17/2021.";
            return texto;
        }
    }
}
