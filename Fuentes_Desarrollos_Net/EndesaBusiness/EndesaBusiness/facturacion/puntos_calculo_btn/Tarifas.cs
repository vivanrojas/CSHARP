using EndesaBusiness.global;
using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types;

namespace EndesaBusiness.facturacion.puntos_calculo_btn
{
    public class Tarifas
    {
        Dictionary<string, string> dic_tarifa;
        public Tarifas()
        {
            dic_tarifa = Carga_Tabla_CNMC_Codigos();
        }


        private Dictionary<string, string> Carga_Tabla_CNMC_Codigos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            string tarifa = "";
            string tipo = "";

            Dictionary<string, string> d = new Dictionary<string, string>();
            try
            {
                strSql = "select tarifa, tipo from lpc_btn_tipo_tarifa";                   
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    tarifa = r["tarifa"].ToString();
                    tipo = r["tipo"].ToString();

                    d.Add(tarifa, tipo);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }


        public string GetTipoTarifa(string tarifa)
        {
            string o;
            if (dic_tarifa.TryGetValue(tarifa, out o))
                return o;
            else
                return null;

        }
    }
}
