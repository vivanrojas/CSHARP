using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EndesaBusiness.utilidades
{
    public class PdteWeb : EndesaEntity.global.LTP
    {

        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "MED_PdteWeb");
        
        private Dictionary<string, EndesaEntity.global.LTP> dic;
        public PdteWeb(List<string> lista_cups13)
        {
            dic = CargaPendiente(lista_cups13);
        }

        private Dictionary<string, EndesaEntity.global.LTP> CargaPendiente(List<string> lista_cups13)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            DateTime maxFecha = new DateTime();

            Dictionary<string, EndesaEntity.global.LTP> dic = new Dictionary<string, EndesaEntity.global.LTP>();

            try
            {

                strSql = "select max(F_ULT_MOD) as maxFecha from cm_pendiente_ml";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                    maxFecha = Convert.ToDateTime(r["maxFecha"]);
                db.CloseConnection();

                //strSql = "select a.EmpresaTitular, a.PM, a.CodContrato, b.aaaammPdte, a.Distribuidora, a.Estado, a.Subestado"
                //+ " from cm_pendiente_ml a"
                //+ " inner join (select PM, max(aaaammPdte) aaaammPdte from cm_pendiente_ml where "
                //+ " substr(PM,1,13) in ("
                //+ "'" + lista_cups13[0] + "'";

                //for (int i = 1; i < lista_cups13.Count; i++)
                //    strSql += ",'" + lista_cups13[i] + "'";

                //strSql += ") group by substr(PM,1,13)) b on"
                //+ " b.PM = a.PM and b.aaaammPdte = a.aaaammPdte where"
                //+ " substr(a.PM,1,13) in ("
                //+ "'" + lista_cups13[0] + "'";

                //for (int i = 1; i < lista_cups13.Count; i++)
                //    strSql += ",'" + lista_cups13[i] + "'";

                //strSql += ") group by substr(a.PM,1,13) order by a.aaaammPdte";

                strSql = "select a.EmpresaTitular, a.PM, a.CodContrato, a.aaaammPdte, a.Distribuidora, a.Estado, a.Subestado"
                + " from cm_pendiente_ml a where"
                + " substr(a.PM,1,13) in ("
                + "'" + lista_cups13[0] + "'";

                for (int i = 1; i < lista_cups13.Count; i++)
                    strSql += ",'" + lista_cups13[i] + "'";

                strSql += ") and"
                    + " a.F_ULT_MOD = '" + maxFecha.ToString("yyyy-MM-dd") + "'"                    
                    + "group by substr(a.PM,1,13) order by a.aaaammPdte";
                
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.global.LTP c = new EndesaEntity.global.LTP();
                    c.cups13 = r["PM"].ToString().Substring(0, 13);
                    //if (r["cups22"] != System.DBNull.Value)
                    //    if (r["cups22"].ToString() != "")
                    //        c.cups20 = r["cups22"].ToString().Substring(0, 20);
                    c.aaaammPdte = Convert.ToInt32(r["aaaammPdte"]);
                    c.cod_contrato = r["CodContrato"].ToString();
                    c.distribuidora = r["Distribuidora"].ToString();
                    c.empresa_titular = GetIDEmpresaTitular(r["EmpresaTitular"].ToString());
                    c.estado_ltp = r["Estado"].ToString();
                    c.subestado_ltp = r["SubEstado"].ToString();
                    
                    dic.Add(c.cups13, c);

                }
                db.CloseConnection();
                return dic;
                
                

            }
            catch (Exception e)
            {
                ficheroLog.AddError("CargaPendiente: " + e.Message);
                return null;
            }
        }

        private int GetIDEmpresaTitular(string empresa)
        {
            int id_empresa = 0;

            switch (empresa)
            {
                case "EE Sucursal Portugal":
                    id_empresa = 80;
                    break;
                case "Endesa Energía":
                    id_empresa = 20;
                    break;
                case "Endesa Energía XXI":
                    id_empresa = 70;
                    break;
                default:
                    break;
            }

            return id_empresa;
        }

        public void GetEstado(string cups13)
        {
            EndesaEntity.global.LTP o;
            if(dic.TryGetValue(cups13, out o))
            {
                this.aaaammPdte = o.aaaammPdte;
                this.cod_contrato = o.cod_contrato;
                this.cups13 = o.cups13;
                //this.cups20 = c.cups20;
                this.distribuidora = o.distribuidora;
                this.empresa_titular = o.empresa_titular;
                this.estado_ltp = o.estado_ltp;
            }
            else
            {
                this.estado_ltp = "FACTURADO";
            }
        }

       
    }
}
