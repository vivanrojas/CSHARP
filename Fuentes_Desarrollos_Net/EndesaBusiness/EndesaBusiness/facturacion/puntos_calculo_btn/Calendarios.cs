using EndesaBusiness.servidores;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.puntos_calculo_btn
{
    public class Calendarios
    {
        Dictionary<string, int> dic_calendarios;
        public Calendarios()
        {
            dic_calendarios = CargaCalendarios();
        }

        private Dictionary<string, int> CargaCalendarios()
        {
            string strSql = "";
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;
            string cpe = "";
            int calendario = 0;

            Dictionary<string, int> d = new Dictionary<string, int>();

            try
            {
                strSql = "select DISTINCT p.TX_CPE, R.R51120200, R2.R51000040"
                     + " FROM APL_INVENTARIO_PUNTOS_ACTIVOS p"
                     + " INNER JOIN t_ges_mensajes T ON"
                     + " T.tx_cpe = p.TX_CPE"
                     + " INNER JOIN R51120000 R ON"
                     + " T.CD_ID = R.TX_ID And T.ID_SECUENCIAL = R.NM_SECUENCIAL"
                     + " INNER JOIN R51000000 R2 ON"
                     + " T.CD_ID = R2.TX_ID And T.ID_SECUENCIAL = R2.NM_SECUENCIAL"
                     + " WHERE R.R51120200 IS NOT NULL AND T.TX_PASO = 'P5100' order by R2.R51000040 DESC";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    cpe = r["TX_CPE"].ToString();
                    calendario = Convert.ToInt32(r["R51120200"]);

                    int o;
                    if (!d.TryGetValue(cpe, out o))
                        d.Add(cpe, calendario);
                }
                ora_db.CloseConnection();

                return d;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Calculo_Prefactura_BTN.CargaCalendarios: " + ex.Message);
                
                return null;
            }
        }

        public string GetCalendario( string cpe)
        {
            int codigo = 0;
            string calendario = "";

            try
            {
                if (dic_calendarios.TryGetValue(cpe, out codigo))
                {
                    switch (codigo)
                    {
                        case 10:
                            calendario = "SIN CICLO";
                            break;
                        case 20:
                            calendario = "DIARIO";
                            break;
                        case 30:
                            calendario = "SEMANAL";
                            break;
                        case 40:
                            calendario = "SEMANAL";
                            break;
                        case 50:
                            calendario = "SEMANAL";
                            break;

                    }
                }

                return calendario;

            }
            catch (Exception ex)
            {
                
                return null;
            }

        }
    }
}
