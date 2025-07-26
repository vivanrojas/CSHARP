using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.sigame
{
    class DistribuidorasGasMySQL
    {
        Dictionary<string, EndesaEntity.sigame.GestGas_Equiv_Distr_PT> dic;
        Dictionary<string, string> dic_ES;
        public DistribuidorasGasMySQL()
        {
            dic = Carga();
            dic_ES = CargaES();
        }

        private Dictionary<string, EndesaEntity.sigame.GestGas_Equiv_Distr_PT> Carga()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, EndesaEntity.sigame.GestGas_Equiv_Distr_PT> d 
                = new Dictionary<string, EndesaEntity.sigame.GestGas_Equiv_Distr_PT>();

            try
            {

                strSql = "Select CUPS, Distribuidora from GestGas_Equiv_Distr_PT";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.sigame.GestGas_Equiv_Distr_PT c = new EndesaEntity.sigame.GestGas_Equiv_Distr_PT();

                    if (r["CUPS"] != System.DBNull.Value)
                        c.cups = r["CUPS"].ToString();

                    if (r["Distribuidora"] != System.DBNull.Value)
                        c.distribuidora = r["Distribuidora"].ToString();

                    EndesaEntity.sigame.GestGas_Equiv_Distr_PT o;
                    if (!d.TryGetValue(c.cups, out o))
                        d.Add(c.cups, c);

                }

                db.CloseConnection();
                return d;

            }catch(Exception e)
            {
                return null;
            }
            
        
        }

        private Dictionary<string, string> CargaES()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            string dis_sigame = "";
            string dis_mysql = "";


            try
            {

                strSql = "Select Distrib_SIGAME, Distribuidora from GestGas_Equiv_Distribuidoras";
                db = new MySQLDB(servidores.MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {                    

                    if (r["Distrib_SIGAME"] != System.DBNull.Value)
                        dis_sigame = r["Distrib_SIGAME"].ToString();

                    if (r["Distribuidora"] != System.DBNull.Value)
                        dis_mysql = r["Distribuidora"].ToString();

                    string o;
                    if (!d.TryGetValue(dis_sigame, out o))
                        d.Add(dis_sigame, dis_mysql);

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


        public string GetDistribuidoraES(string distribuidora_SIGAME)
        {
            string o;
            if (dic_ES.TryGetValue(distribuidora_SIGAME, out o))
                return o;
            else
                return "";
        }

        public string GetDistribuidoraPT(string cups)
        {
            EndesaEntity.sigame.GestGas_Equiv_Distr_PT o;
            if (dic.TryGetValue(cups.Substring(0, 6), out o))
                return o.distribuidora;
            else
                return "";
        }
    }
}
