using EndesaBusiness.servidores;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EndesaBusiness.facturacion
{
    class ConceptosFacturacion
    {

        Dictionary<int, EndesaEntity.facturacion.Conceptos> lc =
            new Dictionary<int, EndesaEntity.facturacion.Conceptos>();
        public ConceptosFacturacion()
        {
            string strSql = "";
            MySQLDB db;
            MySqlCommand command;
            MySqlDataReader reader;

            strSql = "Select CONC_SCE, DESCRIPCION_CORTA,"
                + " DESCRIPCION from fo_cf group by CONC_SCE order by CONC_SCE;";
            db = new MySQLDB(MySQLDB.Esquemas.FAC);
            command = new MySqlCommand(strSql, db.con);
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                EndesaEntity.facturacion.Conceptos c = new EndesaEntity.facturacion.Conceptos();
                c.descripcion_corta = reader["DESCRIPCION_CORTA"].ToString();
                c.descripcion = reader["DESCRIPCION"].ToString();
                lc.Add(Convert.ToInt32(reader["CONC_SCE"]), c);
            }
            db.CloseConnection();

        }

        public string Descripcion_Corta(int conc_sce)
        {

            EndesaEntity.facturacion.Conceptos c = new EndesaEntity.facturacion.Conceptos();
            string value = null;

            if (lc.ContainsKey(conc_sce))
            {
                lc.TryGetValue(conc_sce, out c);
                value = c.descripcion_corta == "" ? c.descripcion : c.descripcion_corta;
            }

            return value;
        }
    }
}
