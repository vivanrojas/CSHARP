using EndesaBusiness.servidores;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.contratacion.Redshift
{
    public class ContratosRosetta
    {
        public Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>> dic { get; set; }

        public ContratosRosetta(List<string> lista_cups20)
        {
            dic = Carga(lista_cups20);
        }

        private Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>> Carga(List<string> lista_cups20)
        {
            servidores.RedShiftServer db;
            OdbcCommand command;
            OdbcDataReader r;
            string strSql = "";

            Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>> d =
                new Dictionary<string, List<EndesaEntity.contratacion.PS_AT_Tabla>>();
            try
            {

                if (lista_cups20.Count > 0)
                {



                    db = new servidores.RedShiftServer(RedShiftServer.Entornos.PROD);
                    command = new OdbcCommand(Consulta(lista_cups20), db.con);
                    r = command.ExecuteReader();
                    while (r.Read())
                    {
                        EndesaEntity.contratacion.PS_AT_Tabla c = new EndesaEntity.contratacion.PS_AT_Tabla();

                        if (r["cd_cups_gas_ext"] != System.DBNull.Value)
                            c.cups20 = r["cd_cups_gas_ext"].ToString();

                        if (r["fh_alta_crto"] != System.DBNull.Value)
                            c.fecha_alta_contrato = Convert.ToDateTime(r["fh_alta_crto"]);

                        if (r["cd_tarifa"] != System.DBNull.Value)
                            c.tarifa = r["cd_tarifa"].ToString();

                        List<EndesaEntity.contratacion.PS_AT_Tabla> o;
                        if (!d.TryGetValue(c.cups20, out o))
                        {
                            o = new List<EndesaEntity.contratacion.PS_AT_Tabla>();
                            o.Add(c);
                            d.Add(c.cups20, o);
                        }
                        else
                            o.Add(c);
                    }
                    db.CloseConnection();
                }
                return d;

            }catch (Exception ex)
            {
                return null;
            }
        }

        private string Consulta(List<string> lista_cups_20)
        {
            string strSql = "";
            bool firstOnly = true;

            strSql = "select c.cd_empr, c.de_empr, c.cd_cups_gas_ext, c.cd_estado_crto,"
                + " c.de_estado, c.fh_alta_crto, c.cd_tarifa"  
                + " from ed_owner.t_ed_h_crtos c where c.cd_cups_gas_ext in ";

            foreach (string p in lista_cups_20)
            {
                if (firstOnly)
                {
                    strSql += "('" + p + "'";
                    firstOnly = false;
                }
                else
                    strSql += ",'" + p + "'";
            }

            strSql += ")"
                + " and c.cd_empr  = 'YXI' order by c.cd_cups_gas_ext";

            return strSql;
        }



    }
}
