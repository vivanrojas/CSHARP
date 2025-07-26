using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types;

namespace EndesaBusiness.facturacion.informes_cnmc
{
    public class Facturas_BTN
    {
        public Facturas_BTN()
        {

        }

        public void Completa_Perfil_tarifa_calendario()
        {
            Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> dic = 
                new Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo>();

            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            bool firstOnly = true;
            StringBuilder sb = new StringBuilder();

            EndesaBusiness.compor.Inventario inventario;

            string cupsree = "";

            int registrosLeidos = 0;
            int totalReg = 0;


            try
            {
                Console.WriteLine("Cargando inventario compor");
                inventario = new compor.Inventario();

                strSql = "select CUPSREE from cdc_051_facturas"
                    + " where Empresa = 'BTN-Portugal'"
                    + " group by CUPSREE";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                Console.WriteLine(strSql);
                while(r.Read())
                {
                    cupsree = r["CUPSREE"].ToString();

                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo o;
                    if(!dic.TryGetValue(cupsree, out o))
                    {
                        inventario.Get_Info_Inventario(cupsree);
                        if (inventario.existe)
                        {
                            o = new EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo();
                            o.perfil = inventario.perfil;
                            o.tarifa = inventario.tarifa;
                            o.calendario = inventario.calendario;
                            dic.Add(cupsree, o);
                        }
                        
                    }
                        
                }
                db.CloseConnection();

                foreach(KeyValuePair<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> p in dic)
                {
                    registrosLeidos++;
                    totalReg++;

                    if (firstOnly)
                    {
                        sb.Append("replace into cdc_051_inventario_btn_datos (cups20, perfil,");
                        sb.Append(" tarifa, calendario) values ");
                        firstOnly = false;
                    }

                    sb.Append("('").Append(p.Key).Append("',");
                    sb.Append("'").Append(p.Value.perfil).Append("',");
                    sb.Append("'").Append(p.Value.tarifa).Append("',");
                    sb.Append("'").Append(p.Value.calendario).Append("'),");


                    if (totalReg == 650)
                    {

                        Console.CursorLeft = 0;
                        Console.Write(registrosLeidos.ToString("N0") + "/" + dic.Count.ToString("N0"));
                        firstOnly = true;
                        db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                        command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                        command.ExecuteNonQuery();
                        db.CloseConnection();
                        sb = null;
                        sb = new StringBuilder();
                        totalReg = 0;
                    }

                }

                if (totalReg > 0)
                {

                    Console.CursorLeft = 0;
                    Console.Write(registrosLeidos.ToString("N0") + "/" + dic.Count.ToString("N0"));
                    firstOnly = true;
                    db = new servidores.MySQLDB(servidores.MySQLDB.Esquemas.FAC);
                    command = new MySqlCommand(sb.ToString().Substring(0, sb.Length - 1), db.con);
                    command.ExecuteNonQuery();
                    db.CloseConnection();
                    sb = null;
                    sb = new StringBuilder();
                    totalReg = 0;
                }


            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }


    }
}
