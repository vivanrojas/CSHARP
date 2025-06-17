using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using Org.BouncyCastle.Bcpg;
using Org.BouncyCastle.Crypto.Operators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.medida.Redshift
{
    public class Estados_Curvas
    {
        public List<string> estados_facturados { get; set; }
        public List<string> estados_registrados { get; set; }

        Dictionary<string, EndesaEntity.medida.Estado_Curva> dic;

        public Estados_Curvas()
        {
            dic = new Dictionary<string, EndesaEntity.medida.Estado_Curva>();
            estados_facturados = new List<string>();
            estados_registrados = new List<string>();
            GetAll();

        }

        private void GetAll()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql = "";
                      

            try
            {

                strSql = "SELECT c.origen, c.estado, c.descripcion_estado,"
                    + " c.orden_busqueda, c.clasificacion"
                    + " FROM cc_p_estados_curvas c ORDER BY c.orden_busqueda";

                db = new MySQLDB(MySQLDB.Esquemas.MED);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.medida.Estado_Curva c = new EndesaEntity.medida.Estado_Curva();
                    c.origen = r["origen"].ToString();
                    c.estado = r["estado"].ToString();
                    c.descripcion_estado = r["descripcion_estado"].ToString();
                    c.orden_busqueda = Convert.ToInt32(r["orden_busqueda"]);
                    c.clasificacion = r["clasificacion"].ToString();

                    if (c.clasificacion == "Facturada")
                        estados_facturados.Add(c.estado);

                    if (c.clasificacion == "Registrada")
                        estados_registrados.Add(c.estado);

                    EndesaEntity.medida.Estado_Curva o;
                    if (!dic.TryGetValue(c.estado, out o))
                    {
                        o = new EndesaEntity.medida.Estado_Curva();
                        dic.Add(c.estado, c);
                    }
                }
                db.CloseConnection();
                
            }
            catch(Exception ex)
            {
                
            }

        }

       public string GetDescripcion_Estado_Curva(string codigo_estado)
        {
            string descripcion = "";
            EndesaEntity.medida.Estado_Curva o;
            if (dic.TryGetValue(codigo_estado, out o))
                descripcion = o.descripcion_estado;

            return descripcion;
        }


    }
}
