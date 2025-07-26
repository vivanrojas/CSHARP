using EndesaBusiness.servidores;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.sigame
{
    class UsoGas
    {

        Dictionary<int, List<EndesaEntity.sigame.Uso_Gas>> dic;
        public UsoGas(DateTime fd, DateTime fh)
        {
            dic = Carga(fd, fh);
        }

        private Dictionary<int, List<EndesaEntity.sigame.Uso_Gas>> Carga(DateTime fd, DateTime fh)
        {
            SQLServer db;
            SqlCommand command;
            SqlDataReader r;
            string strSql = "";

            Dictionary<int, List<EndesaEntity.sigame.Uso_Gas>> d =
                new Dictionary<int, List<EndesaEntity.sigame.Uso_Gas>>();

            try
            {
                #region Query
                strSql = "select T_SGM_M_HIST_USO_GAS.ID_PS, T_SGM_M_HIST_USO_GAS.ID_PMEDIDA,"
                    + " T_SGM_M_HIST_USO_GAS.FH_INICIO, T_SGM_M_HIST_USO_GAS.FH_FIN,"
                    + " T_SGM_P_USO_GAS.DE_USO_GAS,T_SGM_M_HIST_USO_GAS.NM_PORC_USO_GAS"
                    + " from T_SGM_M_HIST_USO_GAS"
                    + " inner join T_SGM_P_USO_GAS on"
                    + " T_SGM_P_USO_GAS.CD_USO_GAS = T_SGM_M_HIST_USO_GAS.CD_USO_GAS"
                    + " where"
                    + " (T_SGM_M_HIST_USO_GAS.FH_INICIO <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " (T_SGM_M_HIST_USO_GAS.FH_FIN >= '" + fh.ToString("yyyy-MM-dd") + "' or"
                    + " T_SGM_M_HIST_USO_GAS.FH_FIN is null))"
                    + " order by T_SGM_M_HIST_USO_GAS.ID_PS, T_SGM_M_HIST_USO_GAS.FH_INICIO";
                #endregion

                db = new SQLServer();
                command = new SqlCommand(strSql, db.con);

                SqlDataAdapter da = new SqlDataAdapter(command);
                r = command.ExecuteReader();

                while (r.Read())
                {
                    EndesaEntity.sigame.Uso_Gas c = new EndesaEntity.sigame.Uso_Gas();
                    c.id_ps = Convert.ToInt32(r["ID_PS"]);
                    if (r["ID_PMEDIDA"] != System.DBNull.Value)
                        c.id_pmedida = Convert.ToInt32(r["ID_PMEDIDA"]);
                    if (r["FH_INICIO"] != System.DBNull.Value)
                        c.fd = Convert.ToDateTime(r["FH_INICIO"]);
                    if (r["FH_FIN"] != System.DBNull.Value)
                        c.fh = Convert.ToDateTime(r["FH_FIN"]);
                    else
                        c.fh = DateTime.MaxValue;

                    if (r["DE_USO_GAS"] != System.DBNull.Value)
                        c.uso_gas = r["DE_USO_GAS"].ToString();

                    if (r["NM_PORC_USO_GAS"] != System.DBNull.Value)
                        c.porcentaje_uso_gas = Convert.ToDouble(r["NM_PORC_USO_GAS"]);

                    List<EndesaEntity.sigame.Uso_Gas> o;
                    if (!d.TryGetValue(c.id_ps, out o))
                    {
                        o = new List<EndesaEntity.sigame.Uso_Gas>();
                        o.Add(c);
                        d.Add(c.id_ps, o);
                    }
                    else
                        o.Add(c);

                }
                db.CloseConnection();
                return d;

            }catch(Exception e)
            {
                return null;
            }
        }

        public string Get_Uso_Gas(int id_ps, DateTime fd, DateTime fh)
        {
            string uso_gas = "";
            bool firstOnly = true;
            List<EndesaEntity.sigame.Uso_Gas> o;

            if (dic.TryGetValue(id_ps, out o))
            {
                for(int i = 0; i < o.Count; i++)
                {
                    if (o[i].fd <= fd && o[i].fh >= fh)
                    {
                        if (firstOnly)
                        {
                            if(o[i].id_pmedida > 0)
                                uso_gas += "PM " + o[i].id_pmedida + " " + o[i].uso_gas + ": "
                                + o[i].porcentaje_uso_gas.ToString().Replace(".", ",") + "%";
                            else
                                uso_gas += o[i].uso_gas + ": "
                                + o[i].porcentaje_uso_gas.ToString().Replace(".", ",") + "%";

                            firstOnly = false;
                        }
                        else
                        {
                            if (o[i].id_pmedida > 0)
                                uso_gas += " PM " + o[i].id_pmedida + " " + o[i].uso_gas + ": "
                                + o[i].porcentaje_uso_gas.ToString().Replace(".", ",") + "%";
                            else
                                uso_gas += o[i].uso_gas + ": "
                                + o[i].porcentaje_uso_gas.ToString().Replace(".", ",") + "%";

                        }
                    }
                        
                }
            }
            return uso_gas;

        }

    }
}
