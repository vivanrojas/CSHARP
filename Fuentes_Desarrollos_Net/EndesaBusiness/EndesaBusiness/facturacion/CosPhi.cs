using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class CosPhi
    {
        List<EndesaEntity.facturacion.CosPhi> lista;
        double[] vectorExcesosCosPhiPorPeriodo;
        double[] totalCosPhiPorPeriodo;
        double[] vectorCosPhi;
        DateTime _fd = new DateTime();
        DateTime _fh = new DateTime();

        string[] lista_calculo { get; set; }

        public CosPhi(DateTime fd, DateTime fh)
        {
            _fd = fd;
            _fh = fh;
            lista = Carga(fd, fh);
            vectorExcesosCosPhiPorPeriodo = new double[7];
            totalCosPhiPorPeriodo = new double[7];
            vectorCosPhi = new double[7];
            lista_calculo = new string[7];
        }

        private List<EndesaEntity.facturacion.CosPhi> Carga(DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            List<EndesaEntity.facturacion.CosPhi> l = new List<EndesaEntity.facturacion.CosPhi>();

            try
            {
                strSql = "SELECT fechadesde, fechahasta, valuefrom, valueto, value, unit"
                    + " FROM cont.eer_cosphi where"
                    + " (fechadesde <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fechaHasta >= '" + fh.ToString("yyyy-MM-dd") + "')";

                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.CosPhi c = new EndesaEntity.facturacion.CosPhi();
                    if (r["fechadesde"] != System.DBNull.Value)
                        c.fecha_desde = Convert.ToDateTime(r["fechadesde"]);
                    if (r["fechahasta"] != System.DBNull.Value)
                        c.fecha_hasta = Convert.ToDateTime(r["fechahasta"]);
                    if (r["valuefrom"] != System.DBNull.Value)
                        c.value_from = Convert.ToDouble(r["valuefrom"]);
                    if (r["valueto"] != System.DBNull.Value)
                        c.value_to = Convert.ToDouble(r["valueto"]);
                    if (r["value"] != System.DBNull.Value)
                        c.value = Convert.ToDouble(r["value"]);
                    if (r["unit"] != System.DBNull.Value)
                        c.unit = r["unit"].ToString();

                    l.Add(c);

                }
                db.CloseConnection();
                return l;
            }catch(Exception e)
            {
                return null;
            }
        }

        public double ImporteCosPhi(int numPeriodosTarifarios, double[] energiaActivaPorPeriodo,
                                    double[] energiaReactivaPorPeriodo)
        {
            double total = 0;
            Dictionary<int, double> dic_CosPhi = new Dictionary<int, double>();
            double vCosPhi = 0;
            double energiaActivaPeriodo = 0;
            double energiaReactivaPeriodo = 0;
            double valor_intermedio = 0;
            double v = 0;

            for(int i = 1; i <= numPeriodosTarifarios; i++)
            {

                energiaActivaPeriodo = energiaActivaPorPeriodo[i];
                if(energiaActivaPeriodo > 0)
                {

                    if(_fd >= new DateTime(2022,01,01))
                        energiaReactivaPeriodo = (energiaReactivaPorPeriodo[i] < 0 ? 0 : energiaReactivaPorPeriodo[i]);
                    else
                        energiaReactivaPeriodo = energiaReactivaPorPeriodo[i];

                    vCosPhi = (energiaActivaPeriodo /
                        Math.Sqrt((energiaActivaPeriodo * energiaActivaPeriodo) + (energiaReactivaPeriodo * energiaReactivaPeriodo)));
                    
                    vectorCosPhi[i] = vCosPhi;
                                            
                }
            }

            // No aplicar el CosPhi en el ultimo periodo
            for(int i = 1;  i < numPeriodosTarifarios; i++)
            {
                v = BuscarValorCosPhi(vectorCosPhi[i]);

                energiaActivaPeriodo = energiaActivaPorPeriodo[i];

                if (_fd >= new DateTime(2022, 01, 01))
                    energiaReactivaPeriodo = (energiaReactivaPorPeriodo[i] < 0 ? 0 : energiaReactivaPorPeriodo[i]);
                else
                    energiaReactivaPeriodo = energiaReactivaPorPeriodo[i];

                if (energiaReactivaPeriodo - (0.33 * energiaActivaPeriodo) > 0)
                {
                    vectorExcesosCosPhiPorPeriodo[i] = Math.Round(energiaReactivaPeriodo - (0.33 * energiaActivaPeriodo), 0);
                    totalCosPhiPorPeriodo[i] = Math.Round((vectorExcesosCosPhiPorPeriodo[i] * v), 6);
                    
                    dic_CosPhi.Add(i, Math.Round(valor_intermedio, 6));

                    lista_calculo[i] =
                        "P" + i + ": " + string.Format("{0:#,##0.00}", energiaReactivaPeriodo)
                        + " kVArh x " + string.Format("{0:#,##0.000000}", v) + " Eur/kVArh = "
                        + string.Format("{0:#,##0.00}", totalCosPhiPorPeriodo[i]) + " Eur, "
                        + "cos phi " + vectorExcesosCosPhiPorPeriodo[i];

                    total = total + totalCosPhiPorPeriodo[i];
                }
            }
            
            return total;

        }

        private double BuscarValorCosPhi(double v)
        {
            EndesaEntity.facturacion.CosPhi c = lista.Find(z => z.value_from <= v && z.value_to >= v);
            return c.value;
        }
    }
}

