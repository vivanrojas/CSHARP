using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class PagosPortugal
    {

        public Dictionary<string, List<EndesaEntity.facturacion.ValorConceptoPortugalRAM>> dic;
        public PagosPortugalConceptos ppc;
        public PagosPortugal(List<string> lista_cups20, DateTime fd, DateTime fh)
        {
            ppc = new PagosPortugalConceptos(fd, fh);
            dic = Carga(lista_cups20, fd, fh);
        }

        private Dictionary<string, List<EndesaEntity.facturacion.ValorConceptoPortugalRAM>> Carga(List<string> lista_cups20, DateTime fd, DateTime fh)
        {
            Dictionary<string, List<EndesaEntity.facturacion.ValorConceptoPortugalRAM>> d = 
                new Dictionary<string, List<EndesaEntity.facturacion.ValorConceptoPortugalRAM>>();
            
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;            

            try
            {
                strSql = "SELECT CPE, CodigoConceptoFacturado, Cantidad, Precio, Valor from epf_facturacion"
                    + " where CPE in ('" + lista_cups20[0] + "'";
                for (int i = 1; i < lista_cups20.Count; i++)
                    strSql += ",'" + lista_cups20[i] + "'";

                strSql += ") and CodigoConceptoFacturado in ('" + ppc.lista_conceptos[0] + "'";

                for (int i = 1; i < ppc.lista_conceptos.Count; i++)
                    strSql += ",'" + ppc.lista_conceptos[i] + "'";

                strSql += ") and (FechaDesde >= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FechaHasta <= '" + fh.ToString("yyyy-MM-dd") + "')";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.RAM_PAGOS_POR);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.ValorConceptoPortugalRAM c = 
                        new EndesaEntity.facturacion.ValorConceptoPortugalRAM();

                    c.cpe = r["CPE"].ToString();
                    c.codigo = r["CodigoConceptoFacturado"].ToString();
                    c.precio = Convert.ToDouble(r["Precio"]);
                    c.cantidad = Convert.ToDouble(r["Cantidad"]);
                    c.valor = Convert.ToDouble(r["Valor"]);

                    List<EndesaEntity.facturacion.ValorConceptoPortugalRAM> o;
                    if (!d.TryGetValue(c.cpe, out o))
                    {
                        o = new List<EndesaEntity.facturacion.ValorConceptoPortugalRAM>();
                        o.Add(c);
                        d.Add(c.cpe, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();
                return d;
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public double GetValue(string cpe, string concepto)
        {
            double total = 0;

            List<EndesaEntity.facturacion.ValorConceptoPortugalRAM> o;
            if (dic.TryGetValue(cpe, out o))
            {
                for(int i = 0; i < o.Count(); i++)
                {
                    if (o[i].codigo == concepto)
                        total += o[i].valor;
                }
            }

            return total;
        }
    }
}
