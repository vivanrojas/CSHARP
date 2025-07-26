using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class TamElectricidad
    {
        Dictionary<string, double> dic;
        public TamElectricidad(List<string> lista_cups20)
        {
            dic = Carga(lista_cups20);
        }

        public TamElectricidad()
        {

        }

        public void Crea_Tam_Num_Facturas()
        {
            EndesaBusiness.contratacion.PS_AT ps = new contratacion.PS_AT();
            
            StringBuilder sb = new StringBuilder();
            bool firstOnly = true;
            int numReg = 0;
            int totalReg = 0;

            try
            {
                foreach(KeyValuePair<string, EndesaEntity.contratacion.PS_AT_Tabla> p in ps.dic)
                {
                    numReg++;
                    totalReg++;

                    if (firstOnly)
                    {
                        sb.Append("select f.CCOUNIPS, count(*) num_facturas from fo f where");
                        sb.Append(" f.CCOUNIPS in ('").Append(p.Value.cups13).Append("'");
                        firstOnly = false;
                    }else
                        sb.Append(",'" ).Append(p.Value.cups13).Append("'");


                    if(numReg == 250)
                    {
                        sb.Append(") group by f.CCOUNIPS");
                        Console.WriteLine("Actualizando " + totalReg + " registros...");
                        ActualizaTablaNumFacturas(sb);
                        firstOnly = true;                        
                        sb = null;
                        sb = new StringBuilder();
                        numReg = 0;
                    }
                    
                }

                if (numReg > 0)
                {
                    sb.Append(")");
                    Console.WriteLine("Actualizando " + totalReg + " registros...");
                    ActualizaTablaNumFacturas(sb);
                    firstOnly = true;
                    sb = null;
                    sb = new StringBuilder();
                    numReg = 0;
                }

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void ActualizaTablaNumFacturas(StringBuilder sb)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;            
            StringBuilder sb2 = new StringBuilder();
            bool firstOnly = true;  

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(sb.ToString(), db.con);
            r = command.ExecuteReader();
            while (r.Read())
            {
                if (firstOnly)
                {
                    sb2.Append("REPLACE INTO tam_num_facturas (cups13, num_facturas) values ");
                    firstOnly=false;
                }

                sb2.Append("('").Append(r["CCOUNIPS"].ToString()).Append("',");
                sb2.Append(Convert.ToInt32(r["num_facturas"])).Append("),");

            }
            db.CloseConnection();

            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(sb2.ToString().Substring(0, sb2.Length - 1), db.con);
            command.ExecuteNonQuery();
            db.CloseConnection();

        }



        private Dictionary<string, double> Carga(List<string> lista_cups20)
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string cups20 = "";

            Dictionary<string, double> d = new Dictionary<string, double>();
            try
            {
                strSql = "select t.CUPS20, t.TAM from tam t" 
                    + " INNER JOIN (SELECT CUPS20, MAX(FechaHastaMaxima) FechaHastaMaxima FROM fact.tam GROUP BY CUPS20) AS p ON"
                    + " p.CUPS20 = t.CUPS20 AND"
                    + " p.FechaHastaMaxima = t.FechaHastaMaxima"
                    + " where"
                    + " t.CUPS20 in ('" + lista_cups20[0] + "'";

                for (int i = 1; i < lista_cups20.Count; i++)
                    strSql += ",'" + lista_cups20[i] + "'";

                strSql += ")";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    double o;
                    if (r["CUPS20"] != System.DBNull.Value)
                    {
                        cups20 = r["CUPS20"].ToString();
                        if (!d.TryGetValue(cups20, out o))
                            d.Add(cups20, Convert.ToDouble(r["TAM"]));
                    }

                }
                db.CloseConnection();


                strSql = "SELECT t.cd_cups, t.tam FROM t_ed_h_sap_tam_agora t"
                    + " where"
                    + " t.cd_cups in ('" + lista_cups20[0] + "'";

                for (int i = 1; i < lista_cups20.Count; i++)
                    strSql += ",'" + lista_cups20[i] + "'";

                strSql += ")";

                db = new MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    double o;
                    if (r["cd_cups"] != System.DBNull.Value)
                    {
                        cups20 = r["cd_cups"].ToString();
                        if (!d.TryGetValue(cups20, out o))
                            if(r["tam"] != System.DBNull.Value)
                                d.Add(cups20, Convert.ToDouble(r["tam"]));
                            else
                                d.Add(cups20, 0.0);

                    }

                }
                db.CloseConnection();


                return d;
            }catch(Exception e)
            {
                return null;
            }
        }

        public double GetTamCups20(string cups20)
        {

            double o = 0;
            if (dic.TryGetValue(cups20, out o))
                return o;
            return o;
                
        }
    }
}
