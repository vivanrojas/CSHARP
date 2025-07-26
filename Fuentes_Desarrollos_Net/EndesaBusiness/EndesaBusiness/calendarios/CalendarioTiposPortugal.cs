using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EndesaBusiness.calendarios
{
    public class CalendarioTiposPortugal : EndesaEntity.facturacion.Fact_pt_tipos_calendarios_Tabla
    {
        public Dictionary<string, EndesaEntity.facturacion.Fact_pt_tipos_calendarios_Tabla> dic { get; set; }

        public CalendarioTiposPortugal()
        {
            dic = Carga();
        }

        public int GetID(string descripcion)
        {
            EndesaEntity.facturacion.Fact_pt_tipos_calendarios_Tabla o;
            if (dic.TryGetValue(descripcion, out o))
                return o.calendario_id;
            else
                return 0;

        }

        private Dictionary<string, EndesaEntity.facturacion.Fact_pt_tipos_calendarios_Tabla> Carga()
        {


            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader r;
            string strSql;

            Dictionary<string, EndesaEntity.facturacion.Fact_pt_tipos_calendarios_Tabla> d = new Dictionary<string, EndesaEntity.facturacion.Fact_pt_tipos_calendarios_Tabla>();

            try
            {
                strSql = "Select calendario_id, descripcion from fact_pt_tipos_calendarios";
                db = new MySQLDB(MySQLDB.Esquemas.FAC);
                command = new MySqlCommand(strSql, db.con);
                r = command.ExecuteReader();
                while (r.Read())
                {
                    EndesaEntity.facturacion.Fact_pt_tipos_calendarios_Tabla c = new EndesaEntity.facturacion.Fact_pt_tipos_calendarios_Tabla();
                    c.calendario_id = Convert.ToInt32(r["calendario_id"]);
                    c.descripcion = r["descripcion"].ToString();

                    d.Add(c.descripcion, c);
                }
                db.CloseConnection();
                return d;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,
                "ClaseCalendario - CargaTerritorios",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                return null;

            }
        }
    }
}
