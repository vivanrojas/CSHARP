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
    
    public class PagosPortugalConceptos
    {
        public Dictionary<string, List<string>> dic { get; set; }
        public List<string> lista_conceptos { get; set; }
        public PagosPortugalConceptos(DateTime fd, DateTime fh)
        {
            lista_conceptos = new List<string>();
            dic = Carga(fd, fh);

        }

        private Dictionary<string, List<string>> Carga(DateTime fd, DateTime fh)
        {
            Dictionary<string, List<string>> d = new Dictionary<string, List<string>>();
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            try
            {
                strSql = "SELECT codigo, codigo_concepto_facturado from ag_pt_fact_pam"
                    + " where (fd <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " fh >= '" + fh.ToString("yyyy-MM-dd") + "')";                    
                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {

                    List<string> o;
                    lista_conceptos.Add(r["codigo_concepto_facturado"].ToString());
                    if (!d.TryGetValue(r["codigo"].ToString(), out o))
                    {
                        o = new List<string>();
                        o.Add(r["codigo_concepto_facturado"].ToString());
                        d.Add(r["codigo"].ToString(), o);
                    }
                    else
                        o.Add(r["codigo_concepto_facturado"].ToString());                   

                }

                db.CloseConnection();
                return d;

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message,
                  "PagosPortugalConceptos.Carga",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error);
                return null;
            }
        }        
    }
}
