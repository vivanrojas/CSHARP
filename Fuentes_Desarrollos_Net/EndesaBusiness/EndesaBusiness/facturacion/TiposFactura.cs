using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    public class TiposFactura : EndesaEntity.facturacion.Tipos_Factura_Tabla
    {
        public Dictionary<int, EndesaEntity.facturacion.Tipos_Factura_Tabla> dic { get; set; }
        private int pNumRegistros;

        public TiposFactura()
        {
            this.pNumRegistros = 0;
            dic = new Dictionary<int, EndesaEntity.facturacion.Tipos_Factura_Tabla>();
            this.CargaDatos();
        }

        private void CargaDatos()
        {
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;
            string strSql;

            try
            {
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                strSql = "select * from fo_p_tipos_factura;";
                command = new MySqlCommand(strSql, db.con);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    pNumRegistros++;
                    EndesaEntity.facturacion.Tipos_Factura_Tabla tf = new EndesaEntity.facturacion.Tipos_Factura_Tabla();
                    tf.id_tipo_factura = Convert.ToInt32(reader["id_tipo_factura"]);
                    tf.descripcion = reader["Descripcion"].ToString();
                    dic.Add(tf.id_tipo_factura, tf);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        public void GetPosicionID(int pos)
        {
            EndesaEntity.facturacion.Tipos_Factura_Tabla o;
            if (dic.TryGetValue(pos, out o))
            {
                this.id_tipo_factura = o.id_tipo_factura;
                this.descripcion = o.descripcion;
            }

        }

        public int GetIDFromDescription(string description)
        {
            return dic.Where(z => z.Value.descripcion == description).Select(z => z.Key).Single();
        }

        public int NumRegistros()
        {
            return dic.Count();
        }
    }
}
