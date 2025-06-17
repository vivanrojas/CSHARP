using EndesaBusiness.servidores;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.sigame
{
    class Albaranes_Clientes
    {
        logs.Log ficheroLog = new logs.Log(Environment.CurrentDirectory, "logs", "SIGAME_T_SGM_M_ALBARANES_CLIENTES");
        Dictionary<int, EndesaEntity.sigame.T_SGM_M_ALBARANES_CLIENTES> dic;

        public Albaranes_Clientes()
        {
            dic = Carga();
        }

        private Dictionary<int, EndesaEntity.sigame.T_SGM_M_ALBARANES_CLIENTES> Carga()
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";

            Dictionary<int, EndesaEntity.sigame.T_SGM_M_ALBARANES_CLIENTES> d = 
                new Dictionary<int, EndesaEntity.sigame.T_SGM_M_ALBARANES_CLIENTES>();

            try
            {
                Console.WriteLine("Cargando datos de T_SGM_M_ALBARANES_CLIENTES");
                #region Query
                strSql = "SELECT ID_PS, CD_NOMBRE_ENTRADA from T_SGM_M_ALBARANES_CLIENTES"
                    + " WHERE (FH_INICIO <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' AND"
                    + " (FH_FIN >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' OR FH_FIN IS NULL))"
                    + " ORDER BY ID_PS";
                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {

                    EndesaEntity.sigame.T_SGM_M_ALBARANES_CLIENTES c = new EndesaEntity.sigame.T_SGM_M_ALBARANES_CLIENTES();
                    c.id_ps = Convert.ToInt32(r["ID_PS"]);
                    if (r["CD_NOMBRE_ENTRADA"] != System.DBNull.Value)
                        c.cd_nombre_entrada = r["CD_NOMBRE_ENTRADA"].ToString();


                    EndesaEntity.sigame.T_SGM_M_ALBARANES_CLIENTES o;
                    if (!d.TryGetValue(c.id_ps, out o))
                            d.Add(c.id_ps, c);
                   

                }
                db.CloseConnection();

                return d;
            }
            catch (Exception e)
            {
                ficheroLog.AddError("Carga: " + e.Message);
                return null;
            }
        }

        public string Get_Codigo_SLM(int id_ps)
        {
            EndesaEntity.sigame.T_SGM_M_ALBARANES_CLIENTES o;
            if (dic.TryGetValue(id_ps, out o))
                return o.cd_nombre_entrada;
            else
                return "";
        }

    }
}
