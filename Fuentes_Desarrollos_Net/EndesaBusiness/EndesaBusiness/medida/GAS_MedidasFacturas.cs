using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida
{
    public class GAS_MedidasFacturas : EndesaEntity.medida.GAS_MedidasFacturas
    {
        Dictionary<int, EndesaEntity.medida.GAS_MedidasFacturas> dic;
        public GAS_MedidasFacturas()
        {
            DateTime fd = DateTime.Now; 
            DateTime fh = DateTime.Now;
            dic = Carga(fd, fh);

        }

        private Dictionary<int, EndesaEntity.medida.GAS_MedidasFacturas> Carga (DateTime fd, DateTime fh)
        {
            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<int, EndesaEntity.medida.GAS_MedidasFacturas> d =
                new Dictionary<int, EndesaEntity.medida.GAS_MedidasFacturas>();

            try
            {

                #region Tubo
                strSql = "select a.`IdPto Medida` as ID, a.Mes, a.Comentario, a.Facturacion, a.Medida"
                    + " from med.GestGas_medidasfacturas a inner join"
                    + " (select b.`IdPto Medida`, max(Mes) Mes from med.GestGas_medidasfacturas b group by b.`IdPto Medida`) as b"
                    + " on a.`IdPto Medida` = b.`IdPto Medida` and"
                    + " a.mes = b.mes"
                    + " where a.Mes > date_format(last_day((now() - interval 2 year)),'%Y%m')";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.GAS_MedidasFacturas c = new EndesaEntity.medida.GAS_MedidasFacturas();

                    if (r["ID"] != System.DBNull.Value)
                        c.id_PS = Convert.ToInt32(r["ID"]);

                    if (r["Mes"] != System.DBNull.Value)
                        c.mes = Convert.ToInt32(r["Mes"]);

                    if (r["Comentario"] != System.DBNull.Value)
                        c.comentario = r["Comentario"].ToString();

                    if (r["Medida"] != System.DBNull.Value)
                        c.medida = Convert.ToDateTime(r["Medida"]);


                    c.facturado = (r["Comentario"].ToString().ToUpper() == "OK");


                    EndesaEntity.medida.GAS_MedidasFacturas o;
                    if (!d.TryGetValue(c.id_PS, out o))
                        d.Add(c.id_PS, c);

                }
                db.CloseConnection();


                #endregion

                #region Cisternas



                strSql = "select a.`IdPto Medida` as ID, a.Mes, a.Comentario, a.Medida"
                    + " from med.GestGas_medidasfacturas_Cisternas a inner join"
                    + " (select b.`IdPto Medida`, max(Mes) Mes from med.GestGas_medidasfacturas_Cisternas b group by b.`IdPto Medida`) as b"
                    + " on a.`IdPto Medida` = b.`IdPto Medida` and"
                    + " a.mes = b.mes"
                    + " where a.Mes > date_format(last_day((now() - interval 2 year)),'%Y%m');";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.GAS_MedidasFacturas c = new EndesaEntity.medida.GAS_MedidasFacturas();

                    if (r["ID"] != System.DBNull.Value)
                        c.id_PS = Convert.ToInt32(r["ID"]);

                    if (r["Mes"] != System.DBNull.Value)
                        c.mes = Convert.ToInt32(r["Mes"]);

                    if (r["Medida"] != System.DBNull.Value)
                        c.medida = Convert.ToDateTime(r["Medida"]);


                    if (r["Comentario"] != System.DBNull.Value)
                        c.comentario = r["Comentario"].ToString();

                    //if (r["Facturacion"] != System.DBNull.Value)
                    c.facturado = (r["Comentario"].ToString().ToUpper() == "OK");

                    EndesaEntity.medida.GAS_MedidasFacturas o;
                    if (!d.TryGetValue(c.id_PS, out o))
                        d.Add(c.id_PS, c);
                }
                db.CloseConnection();

                #endregion

                return d;
            }catch(Exception e)
            {
                return null;
            }
        }

        public void GetID_PS(int id_PS)
        {
            EndesaEntity.medida.GAS_MedidasFacturas o;
            if (dic.TryGetValue(id_PS, out o))
            {
                this.existe = true;
                this.id_PS = o.id_PS;
                this.mes = o.mes;
                this.comentario = o.comentario;
                this.facturado = o.facturado;
                this.medida = o.medida;
            }
        }

    }
}
