using EndesaBusiness.global;
using EndesaBusiness.servidores;
using EndesaEntity.facturacion.puntos_calculo_btn;
using MySql.Data.MySqlClient;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion.puntos_calculo_btn
{
    public class Perfiles
    {
        Dictionary<int, string> dic_descripcion_perfil;
        Dictionary<string, int> dic_cups_perfil;


        public Perfiles()
        {
            dic_descripcion_perfil = Carga_Tabla_Descripcion_Perfiles();
            dic_cups_perfil = CargaPerfiles();
        }

        private Dictionary<int, string> Carga_Tabla_Descripcion_Perfiles()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;
            int codigo = 0;            
            string codigo_perfil = "";

            Dictionary<int, string> d = new Dictionary<int, string>();
            try
            {
                strSql = "select codigo, codigo_perfil from lpc_btn_p_perfiles";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    codigo = Convert.ToInt32(r["codigo"]);
                    codigo_perfil = r["codigo_perfil"].ToString();

                    d.Add(codigo, codigo_perfil);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public string GetCodigoPerfil(int codigo)
        {
            string o;
            if (dic_descripcion_perfil.TryGetValue(codigo, out o))
                return o;
            else
                return null;

        }

        public string GetPerfil( string cpe)
        {

            string codigo_perfil = "";
            int codigo = 0;

            try
            {
                if (dic_cups_perfil.TryGetValue(cpe, out codigo))
                {
                    codigo_perfil = this.GetCodigoPerfil(codigo);
                }

                return codigo_perfil;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Calculo_Prefactura_BTN.GetPerfil: " + ex.Message);                
                return null;
            }
        }

        private Dictionary<string, int> CargaPerfiles()
        {
            string strSql = "";
            OracleServer ora_db;
            OracleCommand ora_command;
            OracleDataReader r;
            string cpe = "";
            int perfil = 0;

            Dictionary<string, int> d = new Dictionary<string, int>();

            try
            {
                strSql = "select DISTINCT p.TX_CPE, R.R51100270, T.DESCRIPCION, R2.R51000040"
                     + " FROM APL_INVENTARIO_PUNTOS_ACTIVOS p"
                     + " INNER JOIN t_ges_mensajes T ON"
                     + " T.tx_cpe = p.TX_CPE"
                     + " INNER JOIN R51100000 R ON T.CD_ID = R.TX_ID And T.ID_SECUENCIAL = R.NM_SECUENCIAL"
                     + " INNER JOIN R51000000 R2 ON T.CD_ID = R2.TX_ID And"
                     + " T.ID_SECUENCIAL = R2.NM_SECUENCIAL"
                     + " left join T12520 T ON R.R51100270 = T.CODIGO"
                     + " WHERE R.R51100270 IS NOT NULL AND T.TX_PASO = 'P5100' order by R2.R51000040 DESC";
                ora_db = new OracleServer(OracleServer.Servidores.COMPOR);
                ora_command = new OracleCommand(strSql, ora_db.con);
                r = ora_command.ExecuteReader();
                while (r.Read())
                {
                    cpe = r["TX_CPE"].ToString();
                    perfil = Convert.ToInt32(r["R51100270"]);

                    int o;
                    if (!d.TryGetValue(cpe, out o))
                        d.Add(cpe, perfil);
                }
                ora_db.CloseConnection();

                return d;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Calculo_Prefactura_BTN.CargaPerfiles: " + ex.Message);
                
                return null;
            }
        }
    }
}
