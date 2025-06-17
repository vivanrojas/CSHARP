using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.eer
{
    public class PuntosMedida
    {
        Dictionary<string, List<EndesaEntity.medida.PuntoMedida>> dic;
        public PuntosMedida()
        {
            dic = Carga(DateTime.Now, DateTime.Now);
        }

        private Dictionary<string, List<EndesaEntity.medida.PuntoMedida>> Carga(DateTime fd, DateTime fh)
        {

            string strSql;
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;

            Dictionary<string, List<EndesaEntity.medida.PuntoMedida>> d
                = new Dictionary<string, List<EndesaEntity.medida.PuntoMedida>>();

            try
            {

                strSql = "SELECT cups22, FechaVigor, FechaAlta, FechaBaja, TipoPM, ModoLectura," 
                    + " Funcion, DireccionEnlace, DireccionPuntoMedida, TelefonoTelemedida, "
                    + " Clave, EstadoTelefono, MarcaMedidaConPerdidas, f_ult_mod"
                    + " FROM cont.eer_puntos_medida m where"
                    + " (FechaAlta <= '" + fd.ToString("yyyy-MM-dd") + "' and"
                    + " FechaBaja >= '" + fh.ToString("yyyy-MM-dd") + "') and"
                    + " Funcion = 'P'";


                db = new MySQLDB(MySQLDB.Esquemas.CON);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.PuntoMedida c = new EndesaEntity.medida.PuntoMedida();
                    if (r["cups22"] != System.DBNull.Value)
                        c.cups22 = r["cups22"].ToString();

                    if (r["FechaVigor"] != System.DBNull.Value)
                        c.fechaVigor = Convert.ToDateTime(r["FechaVigor"]);

                    if (r["FechaAlta"] != System.DBNull.Value)
                        c.fechaAlta = Convert.ToDateTime(r["FechaAlta"]);

                    if (r["FechaBaja"] != System.DBNull.Value)
                        c.fechaBaja = Convert.ToDateTime(r["FechaBaja"]);

                    if (r["TipoPM"] != System.DBNull.Value)
                        c.tipoPuntoMedida = Convert.ToInt32(r["TipoPM"]);

                    if (r["ModoLectura"] != System.DBNull.Value)
                        c.modoLectura = Convert.ToInt32(r["ModoLectura"]);

                    if (r["Funcion"] != System.DBNull.Value)
                        c.funcion = r["Funcion"].ToString();

                    if (r["DireccionEnlace"] != System.DBNull.Value)
                        c.direccion_enlace = Convert.ToInt32(r["DireccionEnlace"]);

                    if (r["DireccionPuntoMedida"] != System.DBNull.Value)
                        c.direccion_enlace = Convert.ToInt32(r["DireccionPuntoMedida"]);

                    if (r["Clave"] != System.DBNull.Value)
                        c.direccion_enlace = Convert.ToInt32(r["Clave"]);

                    if (r["EstadoTelefono"] != System.DBNull.Value)
                        c.direccion_enlace = Convert.ToInt32(r["EstadoTelefono"]);

                    if (r["MarcaMedidaConPerdidas"] != System.DBNull.Value)
                        c.marcaMedidaConPerdidas = r["MarcaMedidaConPerdidas"].ToString() == "S";

                    if (r["DireccionEnlace"] != System.DBNull.Value)
                        c.direccion_enlace = Convert.ToInt32(r["DireccionEnlace"]);

                    List<EndesaEntity.medida.PuntoMedida> l;
                    if (!d.TryGetValue(c.cups22.Substring(0, 20), out l))
                    {
                        l = new List<EndesaEntity.medida.PuntoMedida>();
                        l.Add(c);
                        d.Add(c.cups22.Substring(0, 20), l);
                    }
                    else
                        l.Add(c);                 


                }
                db.CloseConnection();
                return d;

            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
            
        }


        public List<EndesaEntity.medida.PuntoMedida> GetPuntosMedida(string cups20)
        {                       

            List<EndesaEntity.medida.PuntoMedida> o;
            if (dic.TryGetValue(cups20, out o))
                return o;
            else
                return null;
        }
    }
}
