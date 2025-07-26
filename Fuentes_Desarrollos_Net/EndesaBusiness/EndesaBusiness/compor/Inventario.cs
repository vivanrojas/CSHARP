using EndesaBusiness.facturacion.puntos_calculo_btn;
using EndesaBusiness.servidores;
using OfficeOpenXml.ConditionalFormatting;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types;

namespace EndesaBusiness.compor
{
    public class Inventario : EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo
    {
        Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> dic_inventario;
        public Inventario()
        {
            dic_inventario = CargaDatos();
        }

        private Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> CargaDatos()
        {
            

            string strSql = "";
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;

            Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo> d =
                new Dictionary<string, EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo>();

            try
            {

                EndesaBusiness.facturacion.puntos_calculo_btn.Calendarios calendarios =
                   new EndesaBusiness.facturacion.puntos_calculo_btn.Calendarios();

                EndesaBusiness.facturacion.puntos_calculo_btn.Perfiles perfiles =
                    new EndesaBusiness.facturacion.puntos_calculo_btn.Perfiles();

                EndesaBusiness.facturacion.puntos_calculo_btn.Tarifas tarifas =
                    new EndesaBusiness.facturacion.puntos_calculo_btn.Tarifas();



           strSql = "SELECT DISTINCT inv.TX_CPE,"
                     + " NVL(e.TIPO, 'PUNTO NORMAL') AS TIPO,"
                     + " INV.TX_TARIFA_ACCESO"
                     + " FROM APL_INVENTARIO_PUNTOS_ACTIVOS inv"
                     + " LEFT JOIN MED_PUNTOS_ESPECIALES e ON e.CPE = inv.TX_CPE";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo c =
                        new EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo();
                                       

                    if (r["TX_CPE"] != System.DBNull.Value)
                        c.cpe = r["TX_CPE"].ToString();                    

                    if (r["TX_TARIFA_ACCESO"] != System.DBNull.Value)
                        c.tarifa = tarifas.GetTipoTarifa(r["TX_TARIFA_ACCESO"].ToString());

                    if (r["TIPO"] != System.DBNull.Value)
                        c.tipo = r["TIPO"].ToString();

                    
                     c.perfil = perfiles.GetPerfil(c.cpe);                    
                     c.calendario = calendarios.GetCalendario(c.cpe);                   


                    EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo o;
                    if (!d.TryGetValue(c.cpe, out o))
                        d.Add(c.cpe, c);

                }

                ora_db.CloseConnection();

                return d;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Inventario.CargaDatos: " + ex.Message);
                return null;
                
            }

        }
        
        public void Get_Info_Inventario(string cpe)
        {
            EndesaEntity.facturacion.puntos_calculo_btn.puntos_calculo o;
            if (dic_inventario.TryGetValue(cpe, out o))
            {
                this.existe = true;
                this.calendario = o.calendario;
                this.tarifa = o.tarifa;
                this.perfil = o.perfil;
                this.tipo = o.tipo;
            }
        }

    }
}
